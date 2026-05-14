"""
Server-side search functionality for email operations.

This module provides functions for performing server-side searches using
Outlook's AdvancedSearch functionality, which is more efficient for large folders.
"""

# Standard library imports
import time
from datetime import datetime, timedelta, timezone
from typing import Any, List, Optional

# Local application imports
from ..logging_config import get_logger
from ..outlook_session.session_manager import OutlookSessionManager
from .search_common import get_date_limit

logger = get_logger(__name__)

# Default folders to search when cross-folder is requested
DEFAULT_SEARCH_FOLDERS = ["Inbox", "Sent Items"]


def _build_search_criteria(search_term: str, days: int, search_type: str, match_all: bool) -> str:
    """Build SQL search criteria string for Outlook Restrict/AdvancedSearch."""
    date_limit = get_date_limit(days)
    escaped_search_term = search_term.replace("'", "''")
    sql_conditions = []

    sql_conditions.append(
        f"urn:schemas:httpmail:datereceived >= '{date_limit.strftime('%Y-%m-%d')}'"
    )

    if search_type == "subject":
        sql_conditions.append(
            f"urn:schemas:httpmail:subject LIKE '%{escaped_search_term}%'"
        )
    elif search_type == "sender":
        sql_conditions.append(
            f"(urn:schemas:httpmail:fromname LIKE '%{escaped_search_term}%' OR "
            f"urn:schemas:httpmail:senderemail LIKE '%{escaped_search_term}%' OR "
            f"urn:schemas:httpmail:fromemail LIKE '%{escaped_search_term}%')"
        )
    elif search_type == "recipient":
        sql_conditions.append(
            f"urn:schemas:httpmail:to LIKE '%{escaped_search_term}%'"
        )
    elif search_type == "body":
        sql_conditions.append(
            f"urn:schemas:httpmail:textdescription LIKE '%{escaped_search_term}%'"
        )
    elif search_type == "conversation":
        # Search by conversation ID - no date filter needed for thread
        sql_conditions = [
            f"http://schemas.microsoft.com/mapi/proptag/0x3013001F = '{escaped_search_term}'"
        ]

    return "@SQL=" + " AND ".join(sql_conditions)


def _search_single_folder(
    folder, search_criteria: str, max_results: int = 2000, namespace=None
) -> List[Any]:
    """Search a single folder using Items.Restrict, with manual fallback.

    When Restrict returns 0 results (often due to unreliable DAV property
    support e.g. fromname/senderemail), falls back to date-only filter
    + manual property check in Python.
    """
    folder_path = getattr(folder, 'FolderPath', str(folder))
    need_fallback = False
    try:
        items = folder.Items
        restricted_items = items.Restrict(search_criteria)
        results = list(restricted_items)
        logger.info(f"Restrict found {len(results)} results in '{folder.Name}'")
        if results:
            return results[:max_results]
        # 0 results — don't trust it; DAV properties may be unsupported
        need_fallback = True
    except Exception as e:
        logger.warning(f"Restrict failed for '{folder.Name}': {e}")
        need_fallback = True

    if not need_fallback:
        return []

    # Fallback: parse date + search terms from criteria and filter manually.
    # AdvancedSearch is unreliable (disconnected COM objects), so we iterate
    # items by index — the approach proven by find-recent.
    try:
        import re
        date_match = re.search(
            r"datereceived\s*>=\s*'(\d{4}-\d{2}-\d{2})'", search_criteria
        )
        if date_match:
            from datetime import datetime, timezone, timedelta
            date_limit = datetime.strptime(
                date_match.group(1), '%Y-%m-%d'
            ).replace(tzinfo=timezone.utc)
        else:
            date_limit = datetime.now(timezone.utc) - timedelta(days=30)

        term_matches = re.findall(r"LIKE\s+'%([^']+)%'", search_criteria)
        search_terms = [t.lower() for t in term_matches]

        items = folder.Items
        try:
            items.Sort("[ReceivedTime]", True)
        except Exception:
            pass

        total = getattr(items, 'Count', 0)
        results = []

        for i in range(min(total, 2000)):
            try:
                item = items.Item(i + 1)
                if not item:
                    continue
                if (
                    hasattr(item, 'ReceivedTime')
                    and item.ReceivedTime
                ):
                    item_time = item.ReceivedTime
                    if getattr(item_time, 'tzinfo', None) is None:
                        item_time = item_time.replace(tzinfo=timezone.utc)
                    if item_time < date_limit:
                        continue

                if hasattr(item, 'Class') and item.Class != 43:
                    continue

                if search_terms:
                    sender_name = getattr(item, 'SenderName', '') or ''
                    sender_email = (
                        getattr(item, 'SenderEmailAddress', '') or ''
                    )
                    combined = (sender_name + ' ' + sender_email).lower()
                    if not any(term in combined for term in search_terms):
                        continue

                results.append(item)
            except Exception:
                continue

        logger.info(
            f"Manual fallback found {len(results)} results in '{folder.Name}'"
        )
        return results[:max_results]
    except Exception as e2:
        logger.error(
            f"Manual fallback also failed for '{folder.Name}': {e2}"
        )

    return []


def _resolve_folders(session, folder_names: Optional[List[str]] = None) -> List[Any]:
    """Resolve folder names to Outlook folder objects."""
    if not folder_names:
        folder_names = ["Inbox"]

    folders = []
    for name in folder_names:
        try:
            folder = session.get_folder(name)
            if folder:
                folders.append(folder)
            else:
                logger.warning(f"Folder '{name}' not found, skipping")
        except Exception as e:
            logger.warning(f"Error resolving folder '{name}': {e}")

    return folders


def server_side_search(
    folder, search_term: str, days: int, search_type: str, match_all: bool, namespace=None
) -> List[Any]:
    """
    Perform server-side search using Outlook's Restrict method.

    This is more efficient for large folders as it leverages Outlook's indexing.
    """
    try:
        date_limit = get_date_limit(days)
        escaped_search_term = search_term.replace("'", "''")

        sql_conditions = []
        sql_conditions.append(
            f"urn:schemas:httpmail:datereceived >= '{date_limit.strftime('%Y-%m-%d')}'"
        )

        if search_type == "subject":
            sql_conditions.append(
                f"urn:schemas:httpmail:subject LIKE '%{escaped_search_term}%'"
            )
        elif search_type == "sender":
            sql_conditions.append(
                f"(urn:schemas:httpmail:fromname LIKE '%{escaped_search_term}%' OR "
                f"urn:schemas:httpmail:senderemail LIKE '%{escaped_search_term}%' OR "
                f"urn:schemas:httpmail:fromemail LIKE '%{escaped_search_term}%')"
            )
        elif search_type == "recipient":
            sql_conditions.append(
                f"urn:schemas:httpmail:to LIKE '%{escaped_search_term}%'"
            )

        search_criteria = "@SQL=" + " AND ".join(sql_conditions)

        logger.info(f"Server-side search criteria: {search_criteria}")

        folder_path = (
            folder.FolderPath if hasattr(folder, 'FolderPath') else str(folder)
        )
        logger.info(f"Folder path: {folder_path}")

        try:
            items = folder.Items
            restricted_items = items.Restrict(search_criteria)
            results = list(restricted_items)
            logger.info(f"Restrict method completed: found {len(results)} results")
            return results
        except Exception as e:
            logger.warning(f"Restrict method failed: {e}")

            try:
                outlook = (
                    namespace.Application
                    if hasattr(namespace, 'Application')
                    else namespace
                )
                scope = folder_path
                logger.info(f"Using scope: {scope}")

                search_results = outlook.AdvancedSearch(
                    Scope=scope, Filter=search_criteria, SearchSubFolders=True
                )

                max_wait_time = 5
                start_time = time.time()

                while search_results.SearchState != 1:
                    time.sleep(0.1)
                    if time.time() - start_time > max_wait_time:
                        logger.warning("Server-side search timed out")
                        return []

                results = list(search_results.Results)
                logger.info(f"AdvancedSearch completed: found {len(results)} results")
                return results

            except Exception as e2:
                logger.error(f"AdvancedSearch also failed: {e2}")
                return []

        max_wait_time = 5
        start_time = time.time()

        while search_results.SearchState != 1:
            time.sleep(0.1)
            if time.time() - start_time > max_wait_time:
                logger.warning("Server-side search timed out")
                return []

        results = list(search_results.Results)
        logger.info(f"Server-side search completed: found {len(results)} results")
        return results

    except Exception as e:
        logger.error(f"Server-side search failed: {e}")
        logger.error(f"Error type: {type(e)}")
        logger.error(f"Error details: {str(e)}")
        import traceback
        logger.error(f"Traceback: {traceback.format_exc()}")
        return []


def multi_folder_search(
    search_term: str,
    days: int,
    search_type: str,
    folder_names: List[str],
    match_all: bool = True,
) -> List[dict]:
    """
    Search across multiple folders and return combined results as email dicts.

    COM items are extracted to dicts INSIDE the session context so they
    remain usable after the session closes. Each result is tagged with
    its source folder name. Sorted by received time (newest first).

    Args:
        search_term: The term to search for
        days: Number of days to look back
        search_type: Type of search (subject, sender, recipient, body)
        folder_names: List of folder names to search (e.g. ["Inbox", "Sent Items"])
        match_all: Whether to match all terms (AND logic)

    Returns:
        Combined list of email dicts from all folders
    """
    search_criteria = _build_search_criteria(search_term, days, search_type, match_all)
    logger.info(f"Multi-folder search: type={search_type}, folders={folder_names}")

    all_results = []
    seen_ids = set()

    with OutlookSessionManager() as session:
        from .search_common import extract_email_info
        namespace = getattr(session, 'outlook_namespace', None)
        for folder_name in folder_names:
            try:
                folder = session.get_folder(folder_name)
                if not folder:
                    logger.warning(f"Skipping missing folder: {folder_name}")
                    continue

                results = _search_single_folder(folder, search_criteria, namespace=namespace)
                for item in results:
                    entry_id = getattr(item, 'EntryID', '')
                    if entry_id and entry_id not in seen_ids:
                        seen_ids.add(entry_id)
                        # Extract to dict while COM is alive
                        email_data = extract_email_info(item)
                        if email_data:
                            all_results.append(email_data)

                logger.info(f"'{folder_name}': {len(results)} results")
            except Exception as e:
                logger.warning(f"Error searching '{folder_name}': {e}")
                continue

        # Sort by received time, newest first (inside session for COM safety)
        try:
            all_results.sort(
                key=lambda x: x.get("received_time", ""),
                reverse=True,
            )
        except Exception as e:
            logger.warning(f"Error sorting results: {e}")

    logger.info(f"Multi-folder search total: {len(all_results)} results")
    return all_results


def search_by_conversation_id(conversation_id: str, folder_names: Optional[List[str]] = None) -> List[dict]:
    """
    Find all emails sharing the same conversation ID across specified folders.

    Uses Outlook's ConversationID property to track email threads,
    even when subjects change (RE:/Fwd: prefixes, topic changes).
    Returns email dicts extracted inside the session context.

    Args:
        conversation_id: The conversation ID to search for
        folder_names: Folders to search (defaults to Inbox + Sent Items)

    Returns:
        List of email dicts in the same thread, sorted by time
    """
    if not conversation_id:
        return []

    if not folder_names:
        folder_names = ["Inbox", "Sent Items"]

    search_criteria = _build_search_criteria(
        conversation_id, days=365, search_type="conversation", match_all=True
    )
    logger.info(f"Thread search: conversation_id={conversation_id[:30]}...")

    all_results = []
    seen_ids = set()

    with OutlookSessionManager() as session:
        from .search_common import extract_email_info
        namespace = getattr(session, 'outlook_namespace', None)
        for folder_name in folder_names:
            try:
                folder = session.get_folder(folder_name)
                if not folder:
                    continue

                results = _search_single_folder(folder, search_criteria, namespace=namespace)
                for item in results:
                    entry_id = getattr(item, 'EntryID', '')
                    if entry_id and entry_id not in seen_ids:
                        seen_ids.add(entry_id)
                        email_data = extract_email_info(item)
                        if email_data:
                            all_results.append(email_data)
            except Exception as e:
                logger.warning(f"Error in thread search '{folder_name}': {e}")
                continue

        try:
            all_results.sort(
                key=lambda x: x.get("received_time", "")
            )
        except Exception:
            pass

    logger.info(f"Thread search: found {len(all_results)} emails in conversation")
    return all_results


def search_related_emails(
    email_id: str,
    days: int = 90,
    strategies: Optional[List[str]] = None,
) -> dict:
    """
    Find emails related to a given email using multiple strategies.

    Strategies:
    - thread: Same conversation ID (highest confidence)
    - sender: Same sender within expanded time window
    - keyword: Extract key terms and search subject/body

    Args:
        email_id: EntryID of the reference email
        days: Lookback window for sender and keyword strategies
        strategies: List of strategies to use (default: all three)

    Returns:
        Dict with strategy results and combined ranked list
    """
    if strategies is None:
        strategies = ["thread", "sender", "keyword"]

    output = {
        "reference_email": None,
        "thread_results": [],
        "sender_results": [],
        "keyword_results": [],
        "combined": [],
    }

    with OutlookSessionManager() as session:
        # Get the reference email
        try:
            ref_item = session.outlook_namespace.GetItemFromID(email_id)
        except Exception as e:
            logger.error(f"Reference email not found: {e}")
            return output

        from .search_common import extract_email_info

        ref_info = extract_email_info(ref_item)
        output["reference_email"] = ref_info

        seen_ids = {email_id}
        thread_results = []
        sender_results = []
        keyword_results = []

        # Strategy 1: Thread (ConversationID)
        if "thread" in strategies:
            conv_id = ref_info.get("conversation_id", "")
            if conv_id:
                thread_items = search_by_conversation_id(conv_id)
                for item in thread_items:
                    entry_id = item.get('entry_id', '')
                    if entry_id and entry_id not in seen_ids:
                        seen_ids.add(entry_id)
                        # Already a dict from search_by_conversation_id
                        info = dict(item)
                        info["_confidence"] = 1.0
                        info["_strategy"] = "thread"
                        thread_results.append(info)
                logger.info(f"Strategy 'thread': found {len(thread_results)}")
            output["thread_results"] = thread_results

        # Strategy 2: Same sender + time window
        if "sender" in strategies:
            sender_name = ref_info.get("sender", "")
            if sender_name and sender_name != "Unknown":
                sender_criteria = _build_search_criteria(
                    sender_name, days, "sender", match_all=True
                )
                namespace = getattr(session, 'outlook_namespace', None)
                for folder_name in ["Inbox", "Sent Items"]:
                    try:
                        folder = session.get_folder(folder_name)
                        if not folder:
                            continue
                        items = _search_single_folder(folder, sender_criteria, max_results=100, namespace=namespace)
                        for item in items:
                            eid = getattr(item, 'EntryID', '')
                            if eid and eid not in seen_ids:
                                seen_ids.add(eid)
                                info = extract_email_info(item)
                                info["_confidence"] = 0.7
                                info["_strategy"] = "sender"
                                sender_results.append(info)
                    except Exception as e:
                        logger.warning(f"Sender strategy error in '{folder_name}': {e}")
                logger.info(f"Strategy 'sender': found {len(sender_results)}")
            output["sender_results"] = sender_results

        # Strategy 3: Keyword extraction + subject search
        if "keyword" in strategies:
            keywords = _extract_search_keywords(ref_info)
            if keywords:
                namespace = getattr(session, 'outlook_namespace', None)
                for kw in keywords[:3]:  # Limit to top 3 keywords
                    kw_criteria = _build_search_criteria(
                        kw, days, "subject", match_all=True
                    )
                    for folder_name in ["Inbox", "Sent Items"]:
                        try:
                            folder = session.get_folder(folder_name)
                            if not folder:
                                continue
                            items = _search_single_folder(folder, kw_criteria, max_results=50, namespace=namespace)
                            for item in items:
                                eid = getattr(item, 'EntryID', '')
                                if eid and eid not in seen_ids:
                                    seen_ids.add(eid)
                                    info = extract_email_info(item)
                                    info["_confidence"] = 0.4
                                    info["_strategy"] = f"keyword:{kw}"
                                    keyword_results.append(info)
                        except Exception as e:
                            logger.warning(f"Keyword strategy error: {e}")
                logger.info(f"Strategy 'keyword': found {len(keyword_results)}")
            output["keyword_results"] = keyword_results

        # Combine and sort by confidence (desc), then by time (desc)
        all_combined = thread_results + sender_results + keyword_results
        try:
            all_combined.sort(
                key=lambda x: (
                    -x.get("_confidence", 0),
                    x.get("received_time", ""),
                ),
                reverse=False,
            )
            # For equal confidence, newest first
            # Actually let's sort: confidence desc, then time desc
            all_combined.sort(
                key=lambda x: (
                    -x.get("_confidence", 0),
                    x.get("received_time", ""),
                ),
                reverse=False,
            )
            # Re-sort: primary = confidence (high first), secondary = time (new first)
            from operator import itemgetter
            all_combined.sort(
                key=lambda x: x.get("received_time", ""),
                reverse=True,
            )
            all_combined.sort(
                key=lambda x: x.get("_confidence", 0),
                reverse=True,
            )
        except Exception:
            pass

        output["combined"] = all_combined

    return output


def _extract_search_keywords(email_info: dict) -> List[str]:
    """Extract meaningful keywords from email for related search."""
    keywords = []
    subject = email_info.get("subject", "")

    # Remove common prefixes/suffixes
    cleaned = subject
    for prefix in ["RE:", "Re:", "FW:", "Fwd:", "FWD:"]:
        cleaned = cleaned.replace(prefix, "")
    cleaned = cleaned.strip()

    # Extract words longer than 3 chars, excluding common words
    common_words = {
        "the", "and", "for", "from", "this", "that", "with", "have",
        "your", "you", "are", "not", "was", "all", "can", "has",
        "about", "which", "will", "would", "been", "they", "their",
        "please", "thanks", "regards", "dear", "hello", "attached",
    }

    import re
    words = re.findall(r'[A-Za-z0-9_\-\.#]+', cleaned)
    for word in words:
        if len(word) > 3 and word.lower() not in common_words:
            keywords.append(word)

    # Also check for ticket/ID patterns in body
    # (This is lightweight - full body extraction happens at a higher level)

    return list(dict.fromkeys(keywords))[:10]  # Deduplicate, max 10
