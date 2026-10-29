"""
Direct Ribbon Cache Injector for Microsoft Word & PowerPoint
Reverse-engineered from Office WEF AppCommands 18.0 engine.
Forces the native "Gemini Antigravity" button directly onto the Ribbon Home Tab!
"""

import os

WEF_APPCOMMANDS_DIR = os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\AppCommands\18.0")

def inject_ribbon():
    if not os.path.exists(WEF_APPCOMMANDS_DIR):
        print(f"Error: {WEF_APPCOMMANDS_DIR} does not exist.")
        return

    # 1. Word AppCommands definition
    word_cmd_content = (
        "18.0 0 fr-FR WA200099999 WA200099999 wa200099999 1:0:0:1fr-FR 0 https://localhost:18890/index.html "
        "Gemini Gemini Antigravity in Microsoft Office 1:PrimaryCommandSurface:2:2 PrimaryCommandSurface 1:4 TabHome "
        "1:3 GeminiCommandsGroup.Word 1:1 GeminiTaskpaneButton.Word 0:Gemini  Gemini Gemini "
        "3:16:https://localhost:18890/assets/icon-16.png 32:https://localhost:18890/assets/icon-32.png 80:https://localhost:18890/assets/icon-80.png "
        '{"enabled":true,"apioverride":false,"visible":true}   '
        "2:4:2:https://localhost:18890/index.html 3:Office.AutoShowTaskpaneWithDocument 4: 8:Taskpane.Url Gemini   "
        "5:16:https://localhost:18890/assets/icon-16.png 32:https://localhost:18890/assets/icon-32.png 32:https://localhost:18890/assets/icon-32.png "
        '64:https://localhost:18890/assets/icon-64.png 80:https://localhost:18890/assets/icon-80.png {"enabled":true,"apioverride":false,"visible":true}      0:   '
        "0:5:https://localhost:18890/assets/icon-32.png https://localhost:18890/assets/icon-64.png https://localhost:18890/assets/icon-16.png "
        "https://localhost:18890/assets/icon-32.png https://localhost:18890/assets/icon-80.png 0:\n"
    )

    # 2. PowerPoint AppCommands definition
    ppt_cmd_content = (
        "18.0 0 fr-FR WA200099999 WA200099999 wa200099999 1:0:0:1fr-FR 0 https://localhost:18890/index.html "
        "Gemini Gemini Antigravity in Microsoft Office 1:PrimaryCommandSurface:2:2 PrimaryCommandSurface 1:4 TabHome "
        "1:3 GeminiCommandsGroup.PowerPoint 1:1 GeminiTaskpaneButton.PowerPoint 0:Gemini  Gemini Gemini "
        "3:16:https://localhost:18890/assets/icon-16.png 32:https://localhost:18890/assets/icon-32.png 80:https://localhost:18890/assets/icon-80.png "
        '{"enabled":true,"apioverride":false,"visible":true}   '
        "2:4:2:https://localhost:18890/index.html 3:Office.AutoShowTaskpaneWithDocument 4: 8:Taskpane.Url Gemini   "
        "5:16:https://localhost:18890/assets/icon-16.png 32:https://localhost:18890/assets/icon-32.png 32:https://localhost:18890/assets/icon-32.png "
        '64:https://localhost:18890/assets/icon-64.png 80:https://localhost:18890/assets/icon-80.png {"enabled":true,"apioverride":false,"visible":true}      0:   '
        "0:5:https://localhost:18890/assets/icon-32.png https://localhost:18890/assets/icon-64.png https://localhost:18890/assets/icon-16.png "
        "https://localhost:18890/assets/icon-32.png https://localhost:18890/assets/icon-80.png 0:\n"
    )

    # Write command files
    word_cmd_file = os.path.join(WEF_APPCOMMANDS_DIR, "wa20009999918.0Wordfr-FR")
    ppt_cmd_file = os.path.join(WEF_APPCOMMANDS_DIR, "wa20009999918.0PowerPointfr-FR")

    with open(word_cmd_file, "w", encoding="utf-8") as f:
        f.write(word_cmd_content)
    with open(ppt_cmd_file, "w", encoding="utf-8") as f:
        f.write(ppt_cmd_content)

    print(f"[RibbonInject] Wrote Word commands to {word_cmd_file}")
    print(f"[RibbonInject] Wrote PowerPoint commands to {ppt_cmd_file}")

    # 3. Update RibbonCache files
    word_cache_file = os.path.join(WEF_APPCOMMANDS_DIR, "Word.RibbonCache.fr-FR")
    if os.path.exists(word_cache_file):
        with open(word_cache_file, "r", encoding="utf-8") as f:
            cache = f.read()
        if "WA200099999" not in cache:
            cache = cache.strip() + " 1:WA200099999+Omex 1:0 Omex 1:Omex WA200099999 WA200099999 1:0:0:1fr-FR \n"
            with open(word_cache_file, "w", encoding="utf-8") as f:
                f.write(cache)
            print("[RibbonInject] Updated Word.RibbonCache.fr-FR")
        else:
            print("[RibbonInject] WA200099999 already in Word.RibbonCache.fr-FR")

    ppt_cache_file = os.path.join(WEF_APPCOMMANDS_DIR, "PowerPoint.RibbonCache.fr-FR")
    if os.path.exists(ppt_cache_file):
        with open(ppt_cache_file, "r", encoding="utf-8") as f:
            cache = f.read()
        if "WA200099999" not in cache:
            cache = cache.strip() + " 1:WA200099999+Omex 1:0 Omex 1:Omex WA200099999 WA200099999 1:0:0:1fr-FR \n"
            with open(ppt_cache_file, "w", encoding="utf-8") as f:
                f.write(cache)
            print("[RibbonInject] Updated PowerPoint.RibbonCache.fr-FR")
        else:
            print("[RibbonInject] WA200099999 already in PowerPoint.RibbonCache.fr-FR")

if __name__ == "__main__":
    inject_ribbon()
