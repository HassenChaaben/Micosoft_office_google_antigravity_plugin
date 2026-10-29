"""
MCP (Model Context Protocol) & Connectors Bridge
Connects Gemini Antigravity Office to local MCP servers and plugins via real Antigravity CLI queries (`agy mcp list`, `agy plugin list`).
"""

import os
import re
import subprocess
from typing import Dict, Any, List, Optional
from .antigravity_backend import find_agy

class MCPBridge:
    def __init__(self):
        self.gemini_config_dir = os.path.expanduser("~/.gemini/config")
        self.connectors = self.discover_connectors()

    def _query_agy_mcp_list(self, agy_path: str) -> List[Dict[str, Any]]:
        connectors = []
        try:
            res = subprocess.run(
                [agy_path, "mcp", "list"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                encoding="utf-8",
                errors="replace",
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                timeout=10,
            )
            if res.returncode == 0 and res.stdout:
                lines = res.stdout.strip().splitlines()
                # Line 0 is header: NAME TYPE STATUS COMMAND/URL
                for line in lines[1:]:
                    line = line.strip()
                    if not line:
                        continue
                    parts = line.split()
                    if len(parts) >= 3:
                        name = parts[0]
                        mcp_type = parts[1]
                        status = parts[2]
                        cmd_or_url = parts[3] if len(parts) > 3 else ""
                        connectors.append({
                            "id": name,
                            "name": f"MCP: {name}",
                            "status": status,
                            "type": "mcp",
                            "mcp_type": mcp_type,
                            "target": cmd_or_url
                        })
        except Exception as e:
            print(f"[MCPBridge] Error querying agy mcp list: {e}")
        return connectors

    def _query_agy_plugin_list(self, agy_path: str) -> List[Dict[str, Any]]:
        connectors = []
        try:
            res = subprocess.run(
                [agy_path, "plugin", "list"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                encoding="utf-8",
                errors="replace",
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                timeout=10,
            )
            if res.returncode == 0 and res.stdout:
                for line in res.stdout.splitlines():
                    m = re.match(r"^\s*-\s+([a-zA-Z0-9_-]+)(?:\s*\(([^,]+)(?:,\s*([^)]+))?\))?", line)
                    if m:
                        p_name = m.group(1)
                        p_status = m.group(2).strip() if m.group(2) else "enabled"
                        p_ver = m.group(3).strip() if m.group(3) else ""
                        connectors.append({
                            "id": p_name,
                            "name": f"Plugin: {p_name}",
                            "status": p_status,
                            "type": "plugin",
                            "version": p_ver
                        })
        except Exception as e:
            print(f"[MCPBridge] Error querying agy plugin list: {e}")
        return connectors

    def discover_connectors(self) -> List[Dict[str, Any]]:
        active_connectors: List[Dict[str, Any]] = []
        seen_ids = set()

        agy_path = find_agy()
        if agy_path:
            # 1. Real agy mcp list
            for mcp_conn in self._query_agy_mcp_list(agy_path):
                if mcp_conn["id"] not in seen_ids:
                    seen_ids.add(mcp_conn["id"])
                    active_connectors.append(mcp_conn)

            # 2. Real agy plugin list
            for plugin_conn in self._query_agy_plugin_list(agy_path):
                if plugin_conn["id"] not in seen_ids:
                    seen_ids.add(plugin_conn["id"])
                    active_connectors.append(plugin_conn)

        # 3. Fallback scan of ~/.gemini/config/plugins if present
        if os.path.exists(self.gemini_config_dir):
            plugins_dir = os.path.join(self.gemini_config_dir, "plugins")
            if os.path.exists(plugins_dir):
                try:
                    for item in os.listdir(plugins_dir):
                        if item not in seen_ids:
                            p_path = os.path.join(plugins_dir, item)
                            if os.path.isdir(p_path):
                                seen_ids.add(item)
                                active_connectors.append({
                                    "id": item,
                                    "name": f"Plugin: {item}",
                                    "status": "enabled",
                                    "type": "plugin"
                                })
                except Exception:
                    pass

        # 4. Built-in local filesystem connector
        if "fs" not in seen_ids and "filesystem" not in seen_ids:
            seen_ids.add("fs")
            active_connectors.append({
                "id": "fs",
                "name": "Local Filesystem",
                "status": "active",
                "type": "system"
            })

        self.connectors = active_connectors
        return active_connectors

    def get_connector_list(self) -> List[Dict[str, Any]]:
        return self.connectors

    def query_connector(self, connector_id: str, query: str) -> Dict[str, Any]:
        """
        Queries a local MCP connector or plugin for context.
        """
        conn = next((c for c in self.connectors if c["id"] == connector_id), None)
        if conn:
            return {
                "source": connector_id,
                "name": conn.get("name", connector_id),
                "type": conn.get("type", "unknown"),
                "status": conn.get("status", "active"),
                "query": query,
                "summary": f"Connected to {conn.get('name', connector_id)}."
            }
        return {
            "source": connector_id,
            "status": "not_found",
            "query": query,
            "summary": f"Connector '{connector_id}' is not installed or configured."
        }

mcp_bridge = MCPBridge()
