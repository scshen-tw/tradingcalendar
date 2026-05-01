#!/usr/bin/env python3
"""Fetch the latest CBAS mail from Microsoft Graph.

Environment variables:
  GRAPH_CLIENT_ID      Required to enable Graph mode.
  GRAPH_TENANT         Optional, defaults to "consumers" for Outlook.com/Hotmail.
  CBAS_EMAIL_SENDER    Optional sender SMTP address filter.
  CBAS_GRAPH_FOLDER    Optional folder display name, defaults to "cbas".
  CBAS_EMAIL_SUBJECT   Optional subject keyword, defaults to "cb案件整理表".
"""

import json
import os
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

import msal
import requests

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")


GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
SCOPES = ["Mail.Read"]
TOKEN_CACHE = Path(__file__).with_name("graph_token_cache.bin")


@dataclass
class GraphEmail:
    Subject: str
    HTMLBody: str
    ReceivedTime: datetime


def _load_cache():
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE.exists():
        cache.deserialize(TOKEN_CACHE.read_text(encoding="utf-8"))
    return cache


def _save_cache(cache):
    if cache.has_state_changed:
        TOKEN_CACHE.write_text(cache.serialize(), encoding="utf-8")


def _get_token():
    client_id = os.environ.get("GRAPH_CLIENT_ID", "").strip()
    if not client_id:
        print("ℹ️  未設定 GRAPH_CLIENT_ID，略過 Microsoft Graph。")
        return None

    tenant = os.environ.get("GRAPH_TENANT", "consumers").strip() or "consumers"
    cache = _load_cache()
    app = msal.PublicClientApplication(
        client_id=client_id,
        authority=f"https://login.microsoftonline.com/{tenant}",
        token_cache=cache,
    )

    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(f"無法建立 Graph device flow: {flow}")
        print("🔐 Microsoft Graph 需要第一次授權：")
        print(flow["message"])
        result = app.acquire_token_by_device_flow(flow)

    _save_cache(cache)
    if "access_token" not in result:
        raise RuntimeError(f"Graph token 取得失敗: {json.dumps(result, ensure_ascii=False)}")
    return result["access_token"]


def _graph_get(token, path_or_url, params=None):
    url = path_or_url if path_or_url.startswith("https://") else GRAPH_ROOT + path_or_url
    resp = requests.get(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Prefer": 'outlook.body-content-type="html"',
        },
        params=params,
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()


def _iter_folders(token):
    stack = ["/me/mailFolders?$top=100&$select=id,displayName,childFolderCount"]
    while stack:
        url = stack.pop()
        while url:
            data = _graph_get(token, url)
            for folder in data.get("value", []):
                yield folder
                if folder.get("childFolderCount", 0):
                    folder_id = quote(folder["id"], safe="")
                    stack.append(
                        f"/me/mailFolders/{folder_id}/childFolders"
                        "?$top=100&$select=id,displayName,childFolderCount"
                    )
            url = data.get("@odata.nextLink")


def _find_folder(token, folder_name):
    for folder in _iter_folders(token):
        if folder.get("displayName", "").lower() == folder_name.lower():
            return folder
    return None


def fetch_latest_cbas_email():
    token = _get_token()
    if not token:
        return None

    folder_name = os.environ.get("CBAS_GRAPH_FOLDER", "cbas").strip() or "cbas"
    subject_keyword = os.environ.get("CBAS_EMAIL_SUBJECT", "cb案件整理表").strip()
    sender_filter = os.environ.get("CBAS_EMAIL_SENDER", "").strip().lower()

    print(f"🔎 Microsoft Graph 搜尋資料夾「{folder_name}」...")
    folder = _find_folder(token, folder_name)
    if not folder:
        raise RuntimeError(f"Graph 找不到資料夾: {folder_name}")

    print(f"✅ Graph 找到資料夾: {folder.get('displayName')}")
    folder_id = quote(folder["id"], safe="")
    data = _graph_get(
        token,
        f"/me/mailFolders/{folder_id}/messages",
        params={
            "$top": "50",
            "$orderby": "receivedDateTime desc",
            "$select": "subject,receivedDateTime,from,body",
        },
    )

    for msg in data.get("value", []):
        subject = msg.get("subject") or ""
        sender = (
            msg.get("from", {})
            .get("emailAddress", {})
            .get("address", "")
            .lower()
        )
        if subject_keyword.lower() not in subject.lower():
            continue
        if sender_filter and sender_filter != sender:
            continue

        received_raw = msg.get("receivedDateTime", "")
        received = datetime.fromisoformat(received_raw.replace("Z", "+00:00"))
        print(f"✅ Graph 找到郵件: {subject}")
        print(f"   寄件者: {sender or '未知'}")
        print(f"   收信時間: {received}")
        return GraphEmail(
            Subject=subject,
            HTMLBody=msg.get("body", {}).get("content", ""),
            ReceivedTime=received,
        )

    raise RuntimeError(
        f"Graph 找不到主旨含「{subject_keyword}」"
        + (f"、寄件者為 {sender_filter}" if sender_filter else "")
        + " 的郵件"
    )
