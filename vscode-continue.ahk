
; ===== VS Code debug hotkeys sin cambiar el foco (AHK v1) =====
; Este script detecta VS Code estable o Insiders y envía las teclas de depuración.
; ============================================================
; Hotkeys incluidos para depuración en VS Code (sin cambiar foco)
; ------------------------------------------------------------
; Continuar:              Ctrl+Alt+F5        → {F5}
; Paso sobre (Step Over): Ctrl+Alt+F10       → {F10}
; Paso dentro (Step Into):Ctrl+Alt+F11       → {F11}
; Salir (Step Out):       Ctrl+Alt+Shift+F11 → +{F11}
; Reiniciar depuración:   Ctrl+Alt+R         → ^+{F5}
; Detener depuración:     Ctrl+Alt+S         → +{F5}
; Suspender hotkeys:      Ctrl+Alt+0         (toggle)
; Salir del script:       Ctrl+Alt+Esc       (termina la instancia)
; ============================================================

; --- Función para enviar teclas a VS Code (estable o Insiders) ---
SendToVSCode(keys) {
    ; Si existe la estable, usa esa; si no, prueba Insiders
    if WinExist("ahk_exe Code.exe") {
        ControlSend,, %keys%, ahk_exe Code.exe
        return
    }
    if WinExist("ahk_exe Code - Insiders.exe") {
        ControlSend,, %keys%, ahk_exe Code - Insiders.exe
        return
    }
    ; Si no hay VS Code, notifica (opcional)
    TrayTip, VS Code no está abierto, No se enviaron: %keys%, 2, 1
}

; --- Depuración: enviar F5/F10/F11/Shift+F11/Restart/Stop ---
^!F5::SendToVSCode("{F5}")        ; Continuar
^!F10::SendToVSCode("{F10}")      ; Paso sobre (Step Over)
^!F11::SendToVSCode("{F11}")      ; Paso dentro (Step Into)
^!+F11::SendToVSCode("+{F11}")    ; Salir (Step Out)

; Reiniciar y Detener (más fáciles de pulsar que Ctrl+Alt+Shift+F5)
^!R::SendToVSCode("^+{F5}")       ; Reiniciar depuración (Ctrl+Shift+F5)
^!S::SendToVSCode("+{F5}")        ; Detener depuración (Shift+F5)

; --- Gestión del script ---
^!0::Suspend                      ; Suspender/activar los hotkeys (toggle)
^!Esc::ExitApp                    ; Cerrar el script

; (Opcional) confirmación sonora al enviar:
; SoundBeep, 1000, 50