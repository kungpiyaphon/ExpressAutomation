; force_en.ahk (AutoHotkey v2)
LanguageID := "00000409"  ; English US
layout := DllCall("LoadKeyboardLayout", "Str", LanguageID, "UInt", 1)
DllCall("ActivateKeyboardLayout", "Ptr", layout, "UInt", 0)