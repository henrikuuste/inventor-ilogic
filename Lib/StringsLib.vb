' StringsLib.vb - Centralized Estonian UI Strings
' All user-facing text in one place for consistency

Public Module StringsLib
    
    ' ============================================================
    ' DOCUMENT GUARDS
    ' ============================================================
    
    Public Const MSG_NO_ACTIVE_DOCUMENT As String = "Aktiivne dokument puudub."
    Public Const MSG_REQUIRES_ASSEMBLY As String = "See reegel töötab ainult koostuga (.iam)."
    Public Const MSG_REQUIRES_PART As String = "See reegel töötab ainult detailiga (.ipt)."
    Public Const MSG_REQUIRES_DRAWING As String = "See reegel töötab ainult joonisega (.idw)."
    Public Const MSG_REQUIRES_ASSEMBLY_OR_PART As String = "See reegel töötab koostuga (.iam) või detailiga (.ipt)."
    
    ' ============================================================
    ' BUTTONS
    ' ============================================================
    
    Public Const BTN_OK As String = "OK"
    Public Const BTN_CANCEL As String = "Tühista"
    Public Const BTN_APPLY As String = "Rakenda"
    Public Const BTN_CREATE As String = "Loo"
    Public Const BTN_UPDATE As String = "Uuenda"
    Public Const BTN_DELETE As String = "Kustuta"
    Public Const BTN_RUN As String = "Käivita"
    Public Const BTN_CLOSE As String = "Sulge"
    Public Const BTN_SELECT As String = "Vali..."
    Public Const BTN_BROWSE As String = "Sirvi..."
    Public Const BTN_SELECT_ALL As String = "Vali kõik"
    Public Const BTN_SELECT_NONE As String = "Tühjenda"
    Public Const BTN_CLEAR As String = "Tühjenda"
    Public Const BTN_CLEAR_SELECTION As String = "Tühista valik"
    Public Const BTN_PICK_SURFACE As String = "Vali pind"
    Public Const BTN_APPLY_ALL As String = "Rakenda kõigile"
    
    ' ============================================================
    ' PICKER PROMPTS
    ' ============================================================
    
    Public Const PICK_CANCEL_SUFFIX As String = " - ESC tühistamiseks"
    
    Public Const PICK_POINT As String = "Vali punkt" & PICK_CANCEL_SUFFIX
    Public Const PICK_AXIS As String = "Vali telg" & PICK_CANCEL_SUFFIX
    Public Const PICK_PLANE As String = "Vali tasand" & PICK_CANCEL_SUFFIX
    Public Const PICK_FACE As String = "Vali pind" & PICK_CANCEL_SUFFIX
    Public Const PICK_EDGE As String = "Vali serv" & PICK_CANCEL_SUFFIX
    Public Const PICK_OCCURRENCE As String = "Vali komponent" & PICK_CANCEL_SUFFIX
    
    ''' <summary>
    ''' Formats a pick prompt with custom description.
    ''' Example: FormatPickPrompt("Vali alguspunkt") returns "Vali alguspunkt - ESC tühistamiseks"
    ''' </summary>
    Public Function FormatPickPrompt(description As String) As String
        Return description & PICK_CANCEL_SUFFIX
    End Function
    
    ' ============================================================
    ' COMMON LABELS
    ' ============================================================
    
    Public Const LBL_NAME As String = "Nimi:"
    Public Const LBL_DESCRIPTION As String = "Kirjeldus:"
    Public Const LBL_VALUE As String = "Väärtus:"
    Public Const LBL_COUNT As String = "Kogus:"
    Public Const LBL_WIDTH As String = "Laius:"
    Public Const LBL_HEIGHT As String = "Kõrgus:"
    Public Const LBL_THICKNESS As String = "Paksus:"
    Public Const LBL_MATERIAL As String = "Materjal:"
    Public Const LBL_ORIENTATION As String = "Orientatsioon:"
    Public Const LBL_OFFSET As String = "Nihe:"
    Public Const LBL_CENTER_POINT As String = "Keskpunkt:"
    Public Const LBL_START_POINT As String = "Alguspunkt:"
    Public Const LBL_END_POINT As String = "Lõpp-punkt:"
    Public Const LBL_SCALE As String = "Mõõtkava:"
    Public Const LBL_ANGLE As String = "Nurk:"
    Public Const LBL_DISTANCE As String = "Kaugus:"
    Public Const LBL_FILTER As String = "Filter:"
    Public Const LBL_STATUS As String = "Olek:"
    Public Const LBL_SELECTED As String = "Valitud:"
    
    ' ============================================================
    ' COMMON MESSAGES
    ' ============================================================
    
    Public Const MSG_NO_VIEWS_ON_SHEET As String = "Lehel puuduvad vaated."
    Public Const MSG_NO_SELECTION As String = "Midagi pole valitud."
    Public Const MSG_OPERATION_CANCELLED As String = "Toiming tühistatud."
    Public Const MSG_OPERATION_COMPLETE As String = "Toiming lõpetatud."
    Public Const MSG_OPERATION_FAILED As String = "Toiming ebaõnnestus. Vaata logi akent."
    Public Const MSG_CONFIRM_DELETE As String = "Kas oled kindel, et soovid kustutada?"
    Public Const MSG_VALIDATION_ERROR As String = "Viga"
    Public Const MSG_SAVING As String = "Salvestamine..."
    Public Const MSG_LOADING As String = "Laadimine..."
    Public Const MSG_PROCESSING As String = "Töötlemine..."
    
    ' ============================================================
    ' DRAWING-SPECIFIC
    ' ============================================================
    
    Public Const MSG_NO_PART_REFERENCE As String = "Jooniselt ei leitud viidet detailile."
    Public Const MSG_USE_CREATE_DRAWINGS As String = "Kasuta 'Loo 1:1 joonised' funktsiooni jooniste loomiseks."
    Public Const TITLE_ADD_DIMENSIONS As String = "Lisa mõõdud"
    Public Const TITLE_ADD_VIEWS As String = "Lisa vaated"
    Public Const TITLE_CREATE_DRAWINGS As String = "Loo 1:1 joonised"
    Public Const TITLE_UPDATE_DRAWING As String = "Uuenda 1:1 joonis"
    Public Const TITLE_UPDATE_SHEET_SIZE As String = "Uuenda lehe suurus"
    
    ' ============================================================
    ' ELEMENT RELEASE
    ' ============================================================
    
    Public Const TITLE_ELEMENT_RELEASE As String = "Elementide väljastamine"
    Public Const BTN_ALL_ELEMENTS As String = "Kõik elemendid"
    Public Const BTN_FIRST_ELEMENT As String = "Esimene element"
    Public Const MSG_CONFIRM_RELEASE As String = "Kinnita väljastamine"
    Public Const MSG_RELEASE_COMPLETE As String = "Väljastamine lõpetatud"
    
    ' ============================================================
    ' BOM EXPORT
    ' ============================================================
    
    Public Const TITLE_SELECT_TEMPLATE As String = "Vali BOM mall"
    Public Const TITLE_EXPORT_BOM As String = "Ekspordi BOM"
    Public Const MSG_EXCEL_LOCKED As String = "Excel fail on avatud teises rakenduses. Palun sulge see enne jätkamist."
    
    ' ============================================================
    ' PATTERN (KORDUSED)
    ' ============================================================
    
    Public Const TITLE_CENTER_PATTERN As String = "Kordused keskelt"
    Public Const MSG_PATTERN_CREATED As String = "Muster loodud."
    Public Const MSG_PATTERN_FAILED As String = "Mustri loomine ebaõnnestus. Vaata logi akent."
    Public Const MSG_PATTERN_UPDATED As String = "Muster uuendatud."
    
    ' ============================================================
    ' COMPONENTS
    ' ============================================================
    
    Public Const TITLE_CREATE_PARTS As String = "Loo detailid"
    Public Const TITLE_FLAT_PATTERN_VIEWS As String = "Pinnalaotuse vaated"
    Public Const TITLE_SHEET_METAL As String = "Lehtmetall"
    Public Const TITLE_RESTORE_COLORS As String = "Taasta värvid"
    Public Const TITLE_MATERIAL_APPEARANCE As String = "Määra materjalide välimus"
    
    ' ============================================================
    ' COORDINATES
    ' ============================================================
    
    Public Const TITLE_CREATE_UCS As String = "Koordinaadid - Loo UCS"
    
    ' ============================================================
    ' ASSEMBLY
    ' ============================================================
    
    Public Const TITLE_CREATE_BASE_ELEMENT As String = "Loo aluselement"
    Public Const TITLE_VARIABLES As String = "Muutujad"
    Public Const TITLE_SORT_PARTS As String = "Sorteeri detailid"
    Public Const TITLE_NAME_PARTS As String = "Nimeta detailid"
    
    ' ============================================================
    ' FORMAT HELPERS
    ' ============================================================
    
    ''' <summary>
    ''' Creates a rule-specific dialog title.
    ''' Example: FormatDialogTitle("Lisa mõõdud") for drawing dimensions dialog
    ''' </summary>
    Public Function FormatDialogTitle(ruleName As String, Optional subtitle As String = Nothing) As String
        If subtitle IsNot Nothing Then
            Return ruleName & " — " & subtitle
        End If
        Return ruleName
    End Function
    
    ''' <summary>
    ''' Creates a document guard message for a specific rule.
    ''' </summary>
    Public Function FormatGuardMessage(ruleName As String, message As String) As String
        Return message  ' Message is self-contained; ruleName goes in MessageBox title
    End Function
    
    ''' <summary>
    ''' Formats a count message.
    ''' Example: FormatCount(5, "element") returns "5 elementi"
    ''' </summary>
    Public Function FormatCount(count As Integer, itemName As String) As String
        Return count.ToString() & " " & itemName
    End Function
    
    ''' <summary>
    ''' Formats a selected count message.
    ''' Example: FormatSelectedCount(3) returns "Valitud: 3"
    ''' </summary>
    Public Function FormatSelectedCount(count As Integer) As String
        Return LBL_SELECTED & " " & count.ToString()
    End Function
    
End Module
