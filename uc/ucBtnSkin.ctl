VERSION 5.00
Begin VB.UserControl ucBtnSkin 
   AutoRedraw      =   -1  'True
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1305
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   87
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   0
   End
End
Attribute VB_Name = "ucBtnSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'===================================================================================================
'///////////////////////////////////////////////////////////////////////////////////////////////////
'===================================================================================================

' >> Descripción : Control de usuario para usar botones con Skins varios, Botones normales, y de texto con formato
' >> Autor       : Luciano Lodola - http://www.recursosvisualbasic.com.ar/
'
' >> Créditos    : A Gonchuki ( CHAMELEON BUTTON ) por la función para dibujar el FocusRect con el Api DrawFocusRect, y por el código de los eventos  de teclas ( KeyPress, Keydown etc ...)
'                : A Leandro Ascierto por la función de la clase cClassToolBar para dibujar con stretchBlt y bitBlt el bitmap de Skin en el UC

'===================================================================================================
'///////////////////////////////////////////////////////////////////////////////////////////////////
'===================================================================================================



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'===============================================================================================================
' Constantes
'===============================================================================================================
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Const DT_LEFT As Long = &H0
Private Const DT_CALCRECT = &H400
Private Const DT_WORDBREAK = &H10
Private Const DT_CENTER = &H1 Or DT_WORDBREAK Or &H4

Private Const mDefFontNameFormatText    As String = "Verdana"
Private Const mDefFontSizeFormatText    As Integer = 8

' Valores de constantes para los skins. Si se cambia las dimensiones del gráfico, se deben cambiar las dimensiones de estas constantes
Private Const SkinBtnBorderW = 5
Private Const SkinBtnBorderH = 5
Private Const SkinBtnWidth = 15                         ' Ancho del skin
Private Const SkinBtnHeight = 23                        ' alto del Skin

Private Const RGN_COPY As Long = &H5&

' Cantidad de Skins a leer desde el archivo de recursos ( si se agregan mas al archivo .res , indicarlo en esta constante)
Private Const COUNT_SKIN_RES As Long = 21

' Colores por defecto
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const mDefBackColor                 As Long = &H8000000F
Private Const mDefBackColorOver             As Long = &H8000000F
Private Const mDefBackColorDown             As Long = &H8000000F
Private Const mDefBackColorCheck            As Long = &H8000000F
Private Const mDefBackColorDisabled         As Long = &H8000000F

Private Const mDefBorderColorNormal         As Long = vbBlack
Private Const mDefBorderColorOver           As Long = vbBlack
Private Const mDefBorderColorDown           As Long = vbBlack
Private Const mDefBorderColorCheck          As Long = vbWhite
Private Const mDefBorderColorDisabled       As Long = vbBlack

Private Const mDefForeColorNormal           As Long = &H80000012
Private Const mDefForeColorUp               As Long = vbBlue
Private Const mDefForeColorDown             As Long = &H80000012
Private Const mDefForeColorDisabled         As Long = &H80000012
Private Const mDefForeColorCheck            As Long = vbBlue


Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
  
' Constantes para SetWindowLong y GetWindowLong
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
  
' Constantes para SetWindowPos
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOZORDER = &H4
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000

Public bFlagNoUpdateBtn                              As Boolean
Attribute bFlagNoUpdateBtn.VB_VarMemberFlags = "440"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'===============================================================================================================
' Fin de Constantes
'===============================================================================================================
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'===================================================================================================
'///////////////////////////////////////////////////////////////////////////////////////////////////
'===================================================================================================

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'===============================================================================================================
' tipos, enums
'===============================================================================================================
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Enumeración para los Skins
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Enum eSkin
    [VBNet] = 0
    [XPAmarillo] = 1
    [XPBlue] = 2
    [XP] = 3
    [XPRosa] = 4
    [Hotmail] = 5
    [WindowsMedia] = 6
    [WMPCobre] = 7
    [WMPVerde] = 8
    [Word2000] = 9
    [RoyaleXP] = 10
    [WindowsLive] = 11
    [WindowsLive2] = 12
    [GooglePicasa] = 13
    [Office2007BlueTabStrip] = 14
    [Office2007BlueButton] = 15
    [Office2007BlueCheckGreen] = 16
    [Hotmail2] = 17
    [WindowsVista] = 18
    [CBBlue] = 19
    [RoyaleXPWinTaskBar] = 20
    [Vista2] = 21
    [Link] = 22         ' no se usa imagen
    [CustomSkin] = 23   ' skin personalizado
End Enum


'Enum para que indica si se usarán los colores por defecto para cada skin o colores propios
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Enum eColorSchemas
    [anone] = 0     ' sin esquema: si se usa esta opción, el color a usar será el que se indique en la ventana de propiedades, o en tiempo de ejecución
    [useSkins] = 1
End Enum


' Alineación del texto del botón cuando No se usa texto con formato
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Enum eAlign
    [eleft] = 0
    [ecenter] = 1
    [eRight] = 2
End Enum


' Enum para el texto con formato ( Saltos, de líneas, texto, imagen o para dibujar una línea con los métodos gráficos )
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Enum eObject
    [eNewLine] = 0
    [ePicture] = 1
    [eText] = 2
    [eLine] = 3
End Enum


' Type para calcular los datos de fuente para cuando se usa texto con formato
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type tDataDrawText
    lMaxHeightFont                      As Long
    lMaxAscentFont                      As Long
    lWidthText                          As Long
    lDesent                             As Long
End Type


' Estilo de línea y grosor ( para el texto con formato )
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type tLineStyle
    lDrawWidth                          As Long
    lDrawStyle                          As DrawStyleConstants
End Type


' Para almacenar datos cuando se usa texto con formato ( alineación, colores,el texto, propiedades de fuente ...)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type tObjects
    oFont                               As New StdFont
    lForeColor                          As Variant
    sText                               As String
    bSaltoDeLinea                       As Boolean
    lBackColor                          As Variant
    bLine                               As tLineStyle
    Picture                             As StdPicture
    Align                               As eAlign
    TypeObject                          As eObject
End Type
 

' Almacenar valor de cada línea ( para el texto con formato )
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type tLineas
    lMaxHeightFont                      As Long
    lMaxAscentFont                      As Long
    lWidhtLine                          As Long
End Type


'Type para la función del Api GetTextExtentPoint32
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type TEXTMETRIC
    tmHeight                            As Long
    tmAscent                            As Long
    tmDescent                           As Long
    tmInternalLeading                   As Long
    tmExternalLeading                   As Long
    tmAveCharWidth                      As Long
    tmMaxCharWidth                      As Long
    tmWeight                            As Long
    tmOverhang                          As Long
    tmDigitizedAspectX                  As Long
    tmDigitizedAspectY                  As Long
    tmFirstChar                         As Byte
    tmLastChar                          As Byte
    tmDefaultChar                       As Byte
    tmBreakChar                         As Byte
    tmItalic                            As Byte
    tmUnderlined                        As Byte
    tmStruckOut                         As Byte
    tmPitchAndFamily                    As Byte
    tmCharSet                           As Byte
End Type


' para usos varios ( api DrawText y otros )
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type RECT
    Left                                As Long
    Top                                 As Long
    Right                               As Long
    Bottom                              As Long
End Type


' Para usos varios ( Api getCursorPos, WindowFromPoint, otros ..)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type POINTAPI
    X                                   As Long
    Y                                   As Long
End Type


' Enum para pasar a las funciones el estado para los eventos del botón
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Enum eMouseEvent
    [Normal] = 0
    [Up] = 1
    [Down] = 2
End Enum


' Enum para el tipo de botón : Botón Normal, OptionButton o check
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Enum eButtonType
    [eNormalButton] = 0
    [eCheckbox] = 1
    [eOptionbutton] = 2
End Enum


' Enum para la función GetSkinsColors ( Para poder recuperar en tiempo de ejecución los colores de los skin y poder asociarlos a otros controles del programa: Color de fondo del formulario, bordes etc ..)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Enum ePosSkin
    [Normal_Skin] = 70
    [Border_Normal_Skin] = 60
    [Hot_Skin] = 12
    [Border_Hot_Skin] = 0
    [Border_Down_Skin] = 15
    [Down_Skin] = 27
    [Border_Hot_Check_Skin] = 30
    [Hot_Check_Skin] = 42
    [Border_Check_Skin] = 45
    [Check_Skin] = 57
End Enum


' Alineación del texto con formato
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Enum eCaptionAlignment
    [Align_Center] = 0
    [Align_Left] = 1
    [Align_Right] = 2
End Enum


' Estado de los botones
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Enum ButtonState 'de uso interno(inventado)
    [TS_HOT] = 0
    [TS_PRESSED] = 1
    [TS_CHECKED] = 2
    [TS_HOTCHECKED] = 3
    [TS_NORMAL] = 4
    [TS_DISABLED] = 5
    [TS_DISABLED_CHECKED] = 6
End Enum


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'===============================================================================================================
' Fin Enumeraciones y Tipos
'===============================================================================================================
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'===================================================================================================
'///////////////////////////////////////////////////////////////////////////////////////////////////
'===================================================================================================

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'===============================================================================================================
' Declaraciones Apis
'===============================================================================================================
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Sub TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long)
Private Declare Function SelectClipPath Lib "gdi32" (ByVal hdc As Long, ByVal iMode As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
  
Private Const WM_SETREDRAW As Long = &HB&


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'====================================================================================
' Fin de  declaraciones Apis
'====================================================================================
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'===================================================================================================
'///////////////////////////////////////////////////////////////////////////////////////////////////
'===================================================================================================

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'====================================================================================
' Miembros , variables
'====================================================================================
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private mbFocus                             As Boolean
Private hDCSkin                             As Long     ' HDC con el bitmap para el skin Actual
Private LastButton                          As Integer  ' Almacena el último botón presionado
Private LastKeyDown                         As Integer  ' Almacena la útlima tecla

Private lastStat                            As Byte     ' Variable para almacenar el estado del botón ( Normal, presionado ) se usa solo para el evento público Refresh del UC

Private mEnabled                            As Boolean
Private mCaption                            As String
Private mValue                              As Boolean
Private mSkin                               As eSkin
Private mCaptionAlign                       As eCaptionAlignment
Private mButtonType                         As eButtonType
Private mMarginCaption                      As Long
Private mToolTipText                        As String
Private mUseUnderLineMouseUp                As Boolean      ' Subrayar texto en eventos de mouse
Private mUseUnderLineMouseCheck             As Boolean
Private mArrstdPicSkins()                   As StdPicture   ' Array para almacenar los skins que se cargan desde el .Res
Private mEnabledFormatText                  As Boolean      ' para usar o no texto con formato en el botón
Private mSkinEnabled                        As Boolean      ' Habilitar o no el skin

' Colores de fondo
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private mBackColor                          As OLE_COLOR
Private mBackColorOver                      As OLE_COLOR
Private mBackColorDown                      As OLE_COLOR
Private mBackColorCheck                     As OLE_COLOR
Private mBackColorDisabled                  As OLE_COLOR


' Colores ( BackColor y borde ) Para poder recuperar en tiempo de ejecución los colores de los skin y poder asociarlos a otros controles del programa: Color de fondo del formulario, bordes etc ..) .. es opcional
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private mBackColorSkinDefault               As Long
Private mBorderColorSkinDefault             As Long

' Colores de fuente
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private mForeColorNormal                    As OLE_COLOR
Private mForeColorUp                        As OLE_COLOR
Private mForeColorDown                      As OLE_COLOR
Private mForeColorDisabled                  As OLE_COLOR
Private mForeColorCheck                     As OLE_COLOR

' Colores de bordes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private mBorderColorNormal                  As OLE_COLOR
Private mBorderColorOver                    As OLE_COLOR
Private mBorderColorDown                    As OLE_COLOR
Private mBorderColorCheck                   As OLE_COLOR
Private mBorderColorDisabled                As OLE_COLOR



' Vars para el texto con Formato
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private arrLineas()                         As tLineas
Private arrTexts()                          As tObjects
Private arrObjects()                        As tObjects     ' array con los datos para dibujar el texto con formato
Private DataDrawText                        As tDataDrawText
Private mAlignText                          As eAlign
Private mLine                               As tLineStyle
Private mMargin                             As Long
Private mSpacingLine                        As Long
Private mForeColor                          As Long
Private mBackColorFormatText                As Variant
Private mTop                                As Long
Private mColorSchemas                       As eColorSchemas ' Esquema de colores para usar con Skin, o colores personalizados cuando está en 'None'para que se puede indicar cualquier color de fondo, fuente y borde
Private mSkinCustomPicture                  As Picture       ' Skin cargado desde un archivo, .Res, etc ..
Private mUseBackColorContainer              As Boolean           ' Si está en True, al botón con estado normal, se le establece el color de fondo que tenga el contenedor del UC
Private mShowFocusRect                      As Boolean


' Textura
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Para habilitar el uso de texturas
Private mUseTexture                         As Boolean

' Imagen de la textura
Private mPictureTexture                     As StdPicture


' Flgs
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private mbFlagMouseOver                     As Boolean
Private mFlagReadInitProp                   As Boolean ' flag para no redibujar el botón mientras se leen algunas propiedades en el evento ReadProperty
Private mFlagSkinEnabled                    As Boolean


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'===============================================================================================================
' Fin de declaraciones de Variables
'===============================================================================================================
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'===================================================================================================
'///////////////////////////////////////////////////////////////////////////////////////////////////
'===================================================================================================

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'===============================================================================================================
'Declaración de Eventos
'===============================================================================================================
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Event Click()
Public Event MouseOver()
Public Event MouseOut()
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'===============================================================================================================
' Fin de declaración de Eventos
'===============================================================================================================
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'===================================================================================================
'///////////////////////////////////////////////////////////////////////////////////////////////////
'===================================================================================================

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'===============================================================================================================
' Subs, Funciones públicas
'===============================================================================================================
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



' Dibujar el texto con formato
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub FormatTextDraw()
    Dim lWidthLine            As Long
    Dim j                     As Long
    Dim i                     As Long
    Dim lTotalHeight          As Long
    Dim oldDrawWidth          As Integer
    Dim oldDrawStyle          As Integer
    Dim oldForeColor          As Long
    
    If UBound(arrObjects) = 0 Then Exit Sub             ' comprobar que está inicializado el array y que hay texto
                                                            
    With UserControl
        oldDrawWidth = .DrawWidth                            ' guardar valores para luego restaurar al terminar la función
        oldDrawStyle = .DrawStyle
        oldForeColor = .ForeColor
    End With
    
    ReDim arrLineas(0)                                  ' Inicializar el array que contendrá las lineas
    
    For i = 1 To UBound(arrObjects)
        
        Call mFormatTextCalcFont(arrObjects(i))                    ' Calcular valores y datos de la fuente para la palabra actual
        
        With DataDrawText
            ' Comprobar que la linea no se pase en ancho
            If ((lWidthLine + .lWidthText + mMargin) >= (ScaleWidth)) Or _
               (arrObjects(i).bSaltoDeLinea = True) Then
                ' Crear nueva linea
                j = UBound(arrLineas) + 1
                ReDim Preserve arrLineas(j)
                
                ' Guardar los valores máximos ( alto de fuente y el Ascent )
                arrLineas(j).lMaxHeightFont = .lMaxHeightFont
                arrLineas(j).lMaxAscentFont = .lMaxAscentFont
                ' flag para el salto de linea para el otro bucle
                arrObjects(i).bSaltoDeLinea = True
                ' Guardar el ancho de la linea actual
                lWidthLine = .lWidthText
                
            Else
                ' Comprobar si hay nuevos valores máximos para el Ascent y el alto de la fuente
                If arrLineas(j).lMaxAscentFont < .lMaxAscentFont Then
                    arrLineas(j).lMaxAscentFont = .lMaxAscentFont
                End If
                
                If arrLineas(j).lMaxHeightFont < .lMaxHeightFont Then
                    arrLineas(j).lMaxHeightFont = .lMaxHeightFont
                End If
                lWidthLine = lWidthLine + .lWidthText
            End If
            
            ' ancho de linea
            arrLineas(j).lWidhtLine = lWidthLine
            
        End With
    
    Next
    
    For i = LBound(arrLineas) To UBound(arrLineas)
        lTotalHeight = lTotalHeight + arrLineas(i).lMaxHeightFont + mSpacingLine
    Next
    
    ' Dibujar en el HDC
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim x1              As Long
    Dim x2              As Long
    Dim y1              As Long
    Dim y2              As Long
    Dim lPosy           As Long
    Dim lInicio         As Long
    Dim mAlign          As eAlign
    
    Dim lFin            As Long
    Dim r               As RECT
    
    lFin = UBound(arrObjects)
    lInicio = 1
    lPosy = mTop  ' valor inicial de la posición top donde comenzar a dibujar
    
    ' recorrer todas las lineas
    For j = 0 To UBound(arrLineas)
        ' recorrer todos los objetos a dibujar
        For i = lInicio To lFin
            ' Si es una nueva linea, resetear la cordenada x
            If arrObjects(i).bSaltoDeLinea And i <> lInicio Then
               lInicio = i
               x1 = 0
               Exit For
            End If
            
            ' Recuperar datos de la fuente
            Call mFormatTextCalcFont(arrObjects(i))
            
            ' Obtener alineación del texto o de la imagen
            mAlign = arrObjects(i).Align
            
            ' Calcular la posición  x1 y x2 ( Left y ancho ) de acuerdo a la alineación
            Select Case mAlign
                Case ecenter
                    If x1 = 0 Then x1 = ((ScaleWidth - arrLineas(j).lWidhtLine) / 2)
                    x2 = x1 + DataDrawText.lWidthText + x1
                
                Case eRight
                    If x1 = 0 Then x1 = ((ScaleWidth) - (arrLineas(j).lWidhtLine)) - (mMargin / 2)
                    x2 = x1 + DataDrawText.lWidthText
                
                Case eleft
                    If x1 = 0 Then x1 = mMargin / 2
                    x2 = x1 + DataDrawText.lWidthText
            End Select
            
            ' calcular la posición del y1 y el y2 para el rectángulo del dibujo
            y1 = (arrLineas(j).lMaxAscentFont - DataDrawText.lMaxAscentFont) + lPosy
            y2 = (arrLineas(j).lMaxHeightFont) + lPosy
                            
            
            ' Establecer en r, el rectángulo para usar con Drawtext
            Call SetRect(r, x1, y1, x2, y2)
                                    
            ' Para el Resalte del texto
            If Not IsEmpty(arrObjects(i).lBackColor) Then
                Dim RecBrush As RECT, hBrush As Long
                ' valores del rectángulo
                With RecBrush
                    .Bottom = r.Bottom
                    .Left = r.Left
                    .Top = lPosy
                    .Right = x1 + DataDrawText.lWidthText ' ancho de la palabra
                End With
                ' Crear el rectángulo de color
                hBrush = CreateSolidBrush(arrObjects(i).lBackColor)
                ' aplicarlo
                Call FillRect(UserControl.hdc, RecBrush, hBrush)
                Call DeleteObject(hBrush)
                
            End If
            
            Dim lcolor As Long
            If mEnabled Then
               lcolor = arrObjects(i).lForeColor
            Else
               lcolor = mForeColorDisabled
            End If
            
            Select Case arrObjects(i).TypeObject
                
                ' linea
                Case eLine
                                          
                     With UserControl
                        .DrawStyle = arrObjects(i).bLine.lDrawStyle
                        .DrawWidth = arrObjects(i).bLine.lDrawWidth
                     
                         UserControl.Line (x1, y1)-(ScaleWidth - x1, y1), lcolor
                        .ForeColor = lcolor
                     End With
                ' imagen
                Case ePicture
                    
                     ' iconos
                     If arrObjects(i).Picture.Type = 3 Then
                        PaintPicture arrObjects(i).Picture, r.Left, r.Top
                     Else ' bmps, jpgs
                        Dim hDCMemory As Long
                        
                        Dim lWidth As Long
                        Dim lHeight As Long
                        
                        lWidth = ScaleX(arrObjects(i).Picture.Width, vbHimetric, vbPixels)
                        lHeight = ScaleY(arrObjects(i).Picture.Height, vbHimetric, vbPixels)
                          
                        hDCMemory = CreateCompatibleDC(0)
                        
                        Call SelectObject(hDCMemory, arrObjects(i).Picture.Handle)
                        SetStretchBltMode hDCMemory, 4
                     
                        Call TransparentBlt(UserControl.hdc, r.Left, r.Top, lWidth, lHeight, hDCMemory, 0, 0, lWidth, lHeight, GetPixel(hDCMemory, 0, 0))
                     End If
                     
                     Call DeleteDC(hDCMemory)
                     
                ' texto o nueva linea
                Case Else
                     UserControl.ForeColor = lcolor
                     Call DrawText(UserControl.hdc, arrObjects(i).sText, Len(arrObjects(i).sText), r, DT_LEFT)
            End Select
            
            ' si es una linea, resetear la posición x
            If (arrObjects(i).sText = "") Or _
               (arrObjects(i).TypeObject = eLine) Then
               
               x1 = 0
            
            Else
               
               x1 = x1 + DataDrawText.lWidthText
            
            End If
        Next
        ' Obtener la cordenada Y ( Pos y, Alto de la fuente, El espaciado vertical )
        lPosy = lPosy + arrLineas(j).lMaxHeightFont + mSpacingLine
        x1 = 0
    Next
    
    With UserControl
        .DrawWidth = oldDrawWidth
        .DrawStyle = oldDrawStyle
        .ForeColor = oldForeColor
    End With
    
End Sub


' Borrar todo el texto y elementos del texto con formato, y establecer valores por defecto para la fuente y las lineas
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub FormatTextClear()

    ReDim arrObjects(0)
    
    mBackColorFormatText = Empty
    mForeColor = 0
    mAlignText = eleft

    With mLine
        .lDrawStyle = vbSolid
        .lDrawWidth = 1
    End With
    
    'Call FormatTextAdd(" ")

End Sub


' Sub Para establecer el margen del texto, el espaciado entr cada linea y la posición inicial superior donde comenzar a dibujar el texto
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub FormatTextSetup(lMargin As Long, lSpacingLine As Long, lTop As Long)
    mSpacingLine = lSpacingLine
    mMargin = lMargin
    mTop = lTop
End Sub


' Agregar un nuevo salto de linea
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub FormatTextAddNewLine()
    Call mFormatTextSetObject("", eNewLine)
End Sub


' Agregar una nueva imagen
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub FormatTextAddPicture(sPathPicture As String)
        
    On Error Resume Next
    
    Dim img As StdPicture
    
    ' Comprobar que se puede leer la imagen y que el gráfico es válido
    Set img = LoadPicture(sPathPicture)

    If Err.Number = 0 Then
       Set img = Nothing
       ' Agregar imagen
       Call mFormatTextSetObject(sPathPicture, ePicture)
    Else
       MsgBox "No se pudo cargar la imagen: " & sPathPicture, vbCritical, "Error"
    End If
End Sub


' Agregar una nueva Linea para dibujar
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub FormatTextDrawLine( _
        ByVal ColorLine As Long, _
        Optional ByVal StyleLine As DrawStyleConstants, _
        Optional ByVal DrawWidth As Long)
        
    If UBound(arrObjects) <> 0 Then
        ' Salto de linea
        Call FormatTextAddNewLine
        
        ' propiedades de dibujo
        With mLine
            .lDrawStyle = StyleLine
            
            If DrawWidth <= 0 Then
              .lDrawWidth = 1
            Else
              .lDrawWidth = DrawWidth
            End If
            
        End With
        ' añadir la linea
        Call mFormatTextSetObject(ColorLine, eLine)
        ' otro salto de linea para que al agregar el texto siguiente no continue con el final de linea
        Call FormatTextAddNewLine
    End If
End Sub



' Agregar dos lineas
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub FormatTextAddParagraph(AlignmentText As eAlign)

    If UBound(arrObjects) <> 0 Then
       Call mFormatTextSetObject("", eNewLine)
       Call mFormatTextSetObject("", eNewLine)
    End If
    ' Guardar la alineación hasta que no se vuelva a agregar otro párrafo
    mAlignText = AlignmentText
    
End Sub


' Agregar nuevo texto
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub FormatTextAdd( _
    sText As String, _
    Optional ByVal FontName As Variant, _
    Optional ByVal FontSize As Variant, _
    Optional ByVal ForeColor As Variant, _
    Optional ByVal lBackColor As Variant, _
    Optional ByVal Italic As Variant, _
    Optional ByVal Bold As Variant, _
    Optional ByVal UnderLine As Variant, _
    Optional ByVal Strikethru As Variant)
        
    ' Almacenar en un temporal los valores de la fuente
    Dim tForeColor  As Long
    Dim tBackColor  As Variant
    
    tForeColor = mForeColor
    tBackColor = mBackColorFormatText
        
    Dim tempFont As StdFont
    Set tempFont = New StdFont
    
    With tempFont
        .Name = UserControl.Font.Name
        .Bold = UserControl.Font.Bold
        .UnderLine = UserControl.Font.UnderLine
        .Strikethrough = UserControl.Font.Strikethrough
        .Italic = UserControl.Font.Italic
        .Size = UserControl.Font.Size
    End With
    
    ' Cambiar la fuente
    With UserControl.Font
        ' fuente
        If Not IsMissing(FontName) Then .Name = FontName
        If Not IsMissing(FontSize) Then .Size = FontSize
        If Not IsMissing(Bold) Then .Bold = Bold
        If Not IsMissing(Strikethru) Then .Strikethrough = Strikethru
        If Not IsMissing(UnderLine) Then .UnderLine = UnderLine
        If Not IsMissing(Italic) Then .Italic = Italic
        ' color
        If Not IsMissing(ForeColor) Then mForeColor = ForeColor
        If Not IsMissing(lBackColor) Then mBackColorFormatText = lBackColor
        
        
    End With
    
    ' Agregar el texto
    Call mFormatTextAdd(sText)
    
    ' Restaurar los valores
    mForeColor = tForeColor
    mBackColorFormatText = tBackColor
    
    Set UserControl.Font = tempFont
    Set tempFont = Nothing
    
Exit Sub
error_Sub:
MsgBox Err.Description

End Sub




' Rutina que dibuja los diferentes estados del botón : Estado normal, MouseOver, MouseUp etc ...
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DrawSkin( _
    ByVal lHdc As Long, _
    Optional ByVal lPixelWidth As Long = 0, _
    Optional ByVal lPixelHeight As Long = 0, _
    Optional ButtonState As ButtonState = TS_NORMAL)
    
    With UserControl
        ' Si el estilo elegido es de tipo enlace normal, no dibujar el Skin, salir de la rutina y dibujar el caption desde la función mUpdateBtn
        If (mSkin = Link) And (lHdc = .hdc) And mSkinEnabled Then
           
           Exit Sub
        End If
        ' Si no se pasa el parámetro para el ancho y alto, entonces es para el Usercontrol
        If lPixelWidth = 0 And (lHdc = UserControl.hdc) Then lPixelWidth = .ScaleWidth
        If lPixelHeight = 0 And (lHdc = UserControl.hdc) Then lPixelHeight = .ScaleHeight
        
    End With
    
    ' Si está habilitada la opción de user el estado normal como color de fondo del contenedor, ... salir y no dibujar el skin
    If (mUseBackColorContainer And ButtonState = TS_NORMAL And mCaption <> vbNullString And mEnabled) Then
       
       Call mDrawBackColor(Extender.Container.BackColor)
       
       Exit Sub
    
    End If
    
    ' Si el Skin está deshabilitado, Establecer BackColor y dibujar bordes  ...
    ' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If SkinEnabled = False Then
       If mEnabled Then
          If ButtonState = TS_CHECKED Then
            Call mDrawBackColor(mBackColorCheck)
            Call mDrawRectangle(mBorderColorCheck)
          ElseIf ButtonState = TS_HOT Then
            Call mDrawBackColor(mBackColorOver)
            Call mDrawRectangle(mBorderColorOver)
          ElseIf ButtonState = TS_PRESSED Then
            Call mDrawBackColor(mBackColorDown)
            Call mDrawRectangle(mBorderColorDown)
          ElseIf ButtonState = TS_HOTCHECKED Then
            Call mDrawBackColor(mBackColorCheck)
            Call mDrawRectangle(mBorderColorOver)
          ElseIf ButtonState = TS_NORMAL Then
            Call mDrawBackColor(mBackColor)
            Call mDrawRectangle(mBorderColorNormal)
          End If
        Else
            Call mDrawBackColor(mBackColorDisabled)
            Call mDrawRectangle(mBorderColorDisabled)
        End If
            
    ElseIf (ButtonState = TS_NORMAL) And (mSkinEnabled = False) Then
        If mEnabled Then
            Call mDrawBackColor(mBackColor)
            UserControl.ForeColor = mBorderColorNormal
        Else
            Call mDrawBackColor(mBackColorDisabled)
            UserControl.ForeColor = mBorderColorDisabled
        End If
        Call Rectangle(UserControl.hdc, 0, 0, lPixelWidth, lPixelHeight)
    
    ' Si el Skin está Habilitado, dibujar la imagen del estado correspondiente ( MouseUp, MouseDown, Checked, normal  etc ...) ...
    ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    Else
        
        Dim dc As Long
        Dim hDCMemory As Long
        Dim hBmp As Long
        Dim SampleLeft As Long
        
        dc = GetDC(0)
        hDCMemory = CreateCompatibleDC(0)
        hBmp = CreateCompatibleBitmap(dc, lPixelWidth, lPixelHeight)
        Call SelectObject(hDCMemory, hBmp)
        SetStretchBltMode hDCMemory, 4
        
        ' calcular la posición X, del la imagen a dibujar
        ' A partir del pixel 0 hasta el 15 es el estado MouseOver -  Posición 15 MouseDown. - Posición 30 Ckeck - posición 45 MouseOver para el Check - Posición 60 Estado normal - Pixel 75 para el deshabilitado
        SampleLeft = (SkinBtnWidth * ButtonState)
        
        Call BitBlt(hDCMemory, 0, 0, SkinBtnBorderW, SkinBtnBorderH, hDCSkin, SampleLeft, 0, vbSrcCopy)
        Call BitBlt(hDCMemory, lPixelWidth - SkinBtnBorderW, 0, SkinBtnBorderW, SkinBtnBorderH, hDCSkin, SampleLeft + SkinBtnWidth - SkinBtnBorderW, 0, vbSrcCopy)
        Call StretchBlt(hDCMemory, SkinBtnBorderW, 0, lPixelWidth - SkinBtnBorderW * 2, SkinBtnBorderH, hDCSkin, SampleLeft + SkinBtnBorderW, 0, SkinBtnWidth - SkinBtnBorderW * 2, SkinBtnBorderH, vbSrcCopy)
        Call StretchBlt(hDCMemory, 0, SkinBtnBorderH, SkinBtnBorderW, lPixelHeight - SkinBtnBorderH * 2, hDCSkin, SampleLeft, SkinBtnBorderH, SkinBtnBorderW, SkinBtnHeight - SkinBtnBorderH * 2, vbSrcCopy)
        Call StretchBlt(hDCMemory, SkinBtnBorderW, SkinBtnBorderH, lPixelWidth - SkinBtnBorderW * 2, lPixelHeight - SkinBtnBorderH * 2, hDCSkin, SampleLeft + SkinBtnBorderW, SkinBtnBorderH, SkinBtnWidth - SkinBtnBorderW * 2, SkinBtnHeight - SkinBtnBorderH * 2, vbSrcCopy)
        Call StretchBlt(hDCMemory, lPixelWidth - SkinBtnBorderW, SkinBtnBorderH, SkinBtnBorderW, lPixelHeight - SkinBtnBorderH * 2, hDCSkin, SampleLeft + SkinBtnWidth - SkinBtnBorderW, SkinBtnBorderH, SkinBtnBorderW, SkinBtnHeight - SkinBtnBorderH * 2, vbSrcCopy)
        Call BitBlt(hDCMemory, 0, lPixelHeight - SkinBtnBorderH, SkinBtnBorderW, SkinBtnBorderH, hDCSkin, SampleLeft, SkinBtnHeight - SkinBtnBorderH, vbSrcCopy)
        Call StretchBlt(hDCMemory, SkinBtnBorderW, lPixelHeight - SkinBtnBorderH, lPixelWidth - SkinBtnBorderW * 2, SkinBtnBorderH, hDCSkin, SampleLeft + SkinBtnBorderW, SkinBtnHeight - SkinBtnBorderH, SkinBtnWidth - SkinBtnBorderW * 2, SkinBtnBorderH, vbSrcCopy)
        Call BitBlt(hDCMemory, lPixelWidth - SkinBtnBorderW, lPixelHeight - SkinBtnBorderH, SkinBtnBorderW, SkinBtnBorderH, hDCSkin, SampleLeft + SkinBtnWidth - SkinBtnBorderW, SkinBtnHeight - SkinBtnBorderH, vbSrcCopy)
        
        'TransparentBlt hdc, x, y, lPixelWidth, lPixelHeight, hDCMemory, 0, 0, lPixelWidth, lPixelHeight, vbMagenta
        Call BitBlt(lHdc, 0, 0, lPixelWidth, lPixelHeight, hDCMemory, 0, 0, vbSrcCopy)
        
        ' Eliminar los DC temporales
        Call DeleteDC(dc)
        Call DeleteDC(hDCMemory)
        Call DeleteObject(hBmp)
        
    End If
End Sub



' Función opcional para poder recuperar en tiempo de ejecución los colores de los skin y poder asociarlos a otros controles del programa: Color de fondo del formulario, bordes etc ..)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function GetSkinsColors(lImgStateButton As ePosSkin) As Long
    Dim Y As Long
    Dim X As Long
        
    ' setear posición 'y' de la imagen donde se obtendrá el pixel
    If (lImgStateButton = Border_Check_Skin) Or _
       (lImgStateButton = Border_Normal_Skin) Or _
       (lImgStateButton = Border_Down_Skin) Or _
       (lImgStateButton = Border_Hot_Check_Skin) Or _
       (lImgStateButton = Border_Hot_Skin) Then
       Y = 22
    Else
       Y = 5
    End If
    
    X = lImgStateButton
    
    ' obtener desde el DC que tiene la imagen ... el color del pixel y retornar el valor a la función
    GetSkinsColors = GetPixel(hDCSkin, X, Y)
End Function

' Refrescar control
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Refresh()
    Call mUpdateBtn(lastStat)
End Sub

Sub SetControlFlatStyle(ByVal lHwnd As Long)
    Dim lStyle As Long
        lStyle = GetWindowLong(lHwnd, GWL_EXSTYLE)
        lStyle = lStyle And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
        Call SetWindowLong(lHwnd, GWL_EXSTYLE, lStyle)
        Call SetWindowPos(lHwnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE)
End Sub
  
' Función para Establecer transparencia a una ventana
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SetTrans(ByVal lHwnd As Long, lValue As Byte) As Long
    
    Dim lStyle As Long
    lStyle = GetWindowLong(lHwnd, GWL_EXSTYLE)
    lStyle = lStyle Or WS_EX_LAYERED
      
    Call SetWindowLong(lHwnd, GWL_EXSTYLE, lStyle)
    Call SetLayeredWindowAttributes(lHwnd, 0, lValue, LWA_ALPHA)

End Function




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'===============================================================================================================
' Fin de Subs y funciones públicas
'===============================================================================================================
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'===============================================================================================================
' Subs, Funciones privadas
'===============================================================================================================
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Sub mShowErrorChangeProp()
    MsgBox "No se puede cambiar este valor si la propiedad 'ColorSchemas' se encuentra en 'useSkins'", vbExclamation
End Sub

Private Sub mShowErrorChangeProp2()
    MsgBox "Si la propiedad SkinEnabled se encuentra en 'true', no se puede cambiar el valor de esta propiedad", vbExclamation
End Sub

Private Sub mShowErrorChangeProp3()
    MsgBox "El Skin estilo 'Link' solo admite cambio en las propiedades de ForeColor ( Colores de fuente ), pero no en los colores de fondo y en los colores para los bordes", vbExclamation
End Sub

' Función para asignar los colores de fuente de los botones cuando tienen Skin y mColorSchemas es igual a UseSkin
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mSetColorSchemas(lSchema As eColorSchemas)
    
    
    ' Salir de la rutina si se está leyendo y cargando las propiedades desde el Propbag
    If mFlagReadInitProp Then Exit Sub
                
    ' Si el botón tiene Skin y no se usa la opción None para colores de fuente propios .....
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If lSchema = useSkins Then
    
        
        ' 0 -  VBNet
        ''''''''''''''''''''''''''''''''''
        If (mSkin = VBNet) Then
            mForeColorCheck = vbBlack
            mForeColorDown = vbBlack
            mForeColorNormal = vbBlack
            mForeColorUp = vbBlack
            mForeColorDisabled = RGB(144, 144, 144)
            mBorderColorSkinDefault = RGB(212, 230, 253)
            mBackColorSkinDefault = RGB(237, 244, 254)
            Exit Sub
        End If
        
        ' 1 - XP amarillo
        ' '''''''''''''''''''''''''''''''''
        
        If (mSkin = XPAmarillo) Then
            mForeColorCheck = vbBlack
            mForeColorDown = mFlagReadInitProp
            mForeColorNormal = vbBlack
            mForeColorUp = vbBlack
            mForeColorDisabled = RGB(147, 157, 17)
            mBorderColorSkinDefault = RGB(239, 242, 159)
            mBackColorSkinDefault = RGB(251, 252, 227)
            
            Exit Sub
        End If
        
        ' 2 - Xp Azul
        ' ''''''''''''''''''''''''''''''''''''
        
        If (mSkin = XPBlue) Then
           
            mForeColorCheck = vbBlack
            mForeColorDown = mFlagReadInitProp
            mForeColorNormal = vbBlack
            mForeColorUp = vbBlack
            mForeColorDisabled = RGB(190, 190, 190)
            mBorderColorSkinDefault = RGB(232, 236, 255)
            mBackColorSkinDefault = RGB(253, 253, 255)
            
            Exit Sub
        End If
        
        ' 3 - XP
        ' '''''''''''''''''''''''''''''''''''''
        
        If (mSkin = XP) Then
            mForeColorCheck = vbBlack
            mForeColorDown = vbBlack
            mForeColorNormal = vbBlack
            mForeColorUp = vbBlack
            mForeColorDisabled = RGB(190, 190, 190)
            mBackColorSkinDefault = RGB(251, 251, 248)
            mBorderColorSkinDefault = RGB(230, 230, 223)
            Exit Sub
        End If
        
        
        ' 4 - XP Rosa
        ' '''''''''''''''''''''''''''''''''''''''
        If (mSkin = XPRosa) Then
           
            mForeColorCheck = vbBlack
            mForeColorDown = mFlagReadInitProp
            mForeColorNormal = vbBlack
            mForeColorUp = vbBlack
            mForeColorDisabled = RGB(255, 166, 255)
            mBackColorSkinDefault = RGB(255, 242, 255)
            mBorderColorSkinDefault = RGB(255, 213, 255)
            Exit Sub
        End If
        
        
        
        ' 5 - Hotmail
        ' '''''''''''''''''''''''''''''''''''''''
        
        If (mSkin = Hotmail) Then
            mForeColorCheck = vbWhite
            mForeColorDown = vbWhite
            mForeColorNormal = vbWhite
            mForeColorUp = vbWhite
            mForeColorDisabled = 0
            mBackColorSkinDefault = RGB(0, 75, 121)
            mBorderColorSkinDefault = RGB(98, 158, 199)
            
            Exit Sub
        End If
        
        
        
        ' 6 - windows Media
        ' ''''''''''''''''''''''''''''''''''''''''
        
        If (mSkin = WindowsMedia) Then
            mForeColorCheck = vbWhite
            mForeColorDown = vbWhite
            mForeColorNormal = vbWhite
            mForeColorUp = vbWhite
            mForeColorDisabled = RGB(190, 190, 190)
            mBackColorSkinDefault = RGB(58, 64, 78)
            mBorderColorSkinDefault = RGB(96, 106, 133)
            
            Exit Sub
        End If
        
        ' 7 - WMPCobre
        ' ''''''''''''''''''''''''''''''''''''''''
        
        If (mSkin = WMPCobre) Then
           
            mForeColorCheck = vbWhite
            mForeColorDown = vbWhite
            mForeColorNormal = vbWhite
            mForeColorUp = vbWhite
            mForeColorDisabled = RGB(170, 170, 170)
            mBackColorSkinDefault = 0
            mBorderColorSkinDefault = RGB(112, 65, 39)
            
            Exit Sub
        End If
                
        
        ' 8 - Wmp Verde
        ' '''''''''''''''''''''''''''''''''''''''''''''
        
         
        If (mSkin = WMPVerde) Then
            mForeColorCheck = vbWhite
            mForeColorDown = vbWhite
            mForeColorNormal = vbWhite
            mForeColorUp = vbWhite
            mForeColorDisabled = RGB(170, 170, 170)
            mBackColorSkinDefault = 0
            mBorderColorSkinDefault = RGB(0, 98, 0)
            Exit Sub
        End If
        
        
        '  9 - Word
        ' '''''''''''''''''''''''''''''''''''''''''''''''
        
        If (mSkin = Word2000) Then
            mForeColorCheck = vbBlack
            mForeColorDown = vbBlack
            mForeColorNormal = vbBlack
            mForeColorUp = vbBlack
            mForeColorDisabled = RGB(190, 190, 190)
            mBackColorSkinDefault = RGB(248, 247, 243)
            mBorderColorSkinDefault = RGB(227, 223, 212)
            Exit Sub
        End If
        
        
        ' 10 - Royale XP
        ' '''''''''''''''''''''''''''''''''''''''''''''''
        
        If (mSkin = RoyaleXP) Then
            mForeColorCheck = vbWhite
            mForeColorDown = vbBlack
            mForeColorNormal = vbBlack
            mForeColorUp = vbBlack
            mForeColorDisabled = RGB(190, 190, 190)
            mBackColorSkinDefault = RGB(251, 252, 255)
            mBorderColorSkinDefault = RGB(214, 213, 217)
            
            Exit Sub
        End If
        
        
        ' 11 -  Windows Live
        ' '''''''''''''''''''''''''''''''''''''''''''''''
        
        If (mSkin = WindowsLive) Then
            mForeColorCheck = vbBlack
            mForeColorDown = vbBlack
            mForeColorNormal = vbBlack
            mForeColorUp = vbBlack
            mForeColorDisabled = RGB(190, 190, 190)
            mBackColorSkinDefault = vbWhite
            mBorderColorSkinDefault = RGB(236, 236, 238)
            Exit Sub
        End If
        
        
        ' 12 -  Windows Live 2
        ' '''''''''''''''''''''''''''''''''''''''''''''''
        
        If (mSkin = WindowsLive2) Then
            mForeColorCheck = vbBlack
            mForeColorDown = vbBlack
            mForeColorNormal = vbBlack
            mForeColorUp = vbBlack
            mForeColorDisabled = RGB(190, 190, 190)
            mBackColorSkinDefault = RGB(243, 243, 243)
            mBorderColorSkinDefault = RGB(227, 227, 227)
            Exit Sub
        End If
    
        
                
        ' 13 - google Picasa
        ' '''''''''''''''''''''''''''''''''''''''''''''''''
        If (mSkin = GooglePicasa) Then
            mForeColorCheck = vbWhite
            mForeColorDown = vbBlack
            mForeColorNormal = vbBlack
            mForeColorUp = vbBlack
            mForeColorDisabled = RGB(180, 180, 180)
            mBackColorSkinDefault = RGB(233, 233, 233)
            mBorderColorSkinDefault = RGB(202, 202, 202)
            
            Exit Sub
        End If
        
        
        ' 14 - Office 2007 Blue Tab strip
        ' ''''''''''''''''''''''''''''''''''''''''''''''''''
        If (mSkin = Office2007BlueTabStrip) Then
           mForeColorCheck = vbWhite
           mForeColorDown = vbBlack
           mForeColorNormal = RGB(34, 66, 125)
           mForeColorUp = vbBlack
           mForeColorDisabled = RGB(101, 146, 214)
           mBackColorSkinDefault = RGB(240, 244, 255)
           mBorderColorSkinDefault = RGB(191, 209, 233)
           
           Exit Sub
        End If
        
        
        ' 15 - Office 2007 Blue botón
        ' ''''''''''''''''''''''''''''''''''''''''''''''''''
        If (mSkin = Office2007BlueButton) Then
           mForeColorCheck = RGB(41, 90, 171)
           mForeColorDown = vbWhite
           mForeColorNormal = RGB(41, 90, 171)
           mForeColorUp = RGB(41, 90, 171)
           mForeColorDisabled = RGB(161, 161, 146)
           mBackColorSkinDefault = RGB(251, 252, 255)
           mBorderColorSkinDefault = RGB(213, 214, 218)
           Exit Sub
        End If
        
        
        
        
        ' 16 - Office 2007 Blue Check Green
        ' ''''''''''''''''''''''''''''''''''''''''''''''''''
        If (mSkin = Office2007BlueCheckGreen) Then
           mForeColorNormal = RGB(23, 51, 96)
           mForeColorCheck = mForeColorNormal
           mForeColorDown = mForeColorNormal
           mForeColorUp = mForeColorNormal
           mForeColorDisabled = RGB(115, 155, 221)
           mBackColorSkinDefault = RGB(240, 244, 255)
           mBorderColorSkinDefault = RGB(191, 209, 233)
           
           Exit Sub
        End If
        
                
        ' 17 - Hotmail 2
        ' '''''''''''''''''''''''''''''''''''''''''''''''''''
        If (mSkin = Hotmail2) Then
            mForeColorCheck = vbBlack
            mForeColorDown = vbWhite
            mForeColorNormal = vbWhite
            mForeColorUp = vbWhite
            mForeColorDisabled = RGB(161, 161, 146)
            mBackColorSkinDefault = RGB(39, 75, 22)
            mBorderColorSkinDefault = RGB(66, 95, 50)
           
           Exit Sub
        End If
        
        
                
        ' 18 - windows Vista
        ' ''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If (mSkin = WindowsVista) Then
            mForeColorCheck = vbBlack
            mForeColorDown = vbBlack
            mForeColorNormal = RGB(48, 98, 139)
            mForeColorUp = vbBlack
            mForeColorDisabled = RGB(161, 161, 146)
            mBackColorSkinDefault = RGB(251, 251, 251)
            mBorderColorSkinDefault = RGB(215, 215, 215)
            
            Exit Sub
        End If
        
        
        ' 19 - CBBlue
        ' '''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If (mSkin = CBBlue) Then
           mForeColorCheck = vbWhite
           mForeColorDown = vbWhite
           mForeColorNormal = vbBlack
           mForeColorUp = vbWhite
           mForeColorDisabled = RGB(111, 153, 200)
           mBackColorSkinDefault = RGB(231, 238, 245)
           mBorderColorSkinDefault = RGB(210, 223, 238)
           
           Exit Sub
        End If
        
        ' 20
        ' '''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If (mSkin = RoyaleXPWinTaskBar) Then
           mForeColorCheck = vbWhite
           mForeColorDown = vbWhite
           mForeColorNormal = vbWhite
           mForeColorUp = vbWhite
           mForeColorDisabled = RGB(180, 180, 180)
           mBackColorSkinDefault = RGB(51, 103, 189)
           mBorderColorSkinDefault = RGB(210, 223, 238)
           Exit Sub
        End If
        
        ' 20
        ' '''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If (mSkin = Vista2) Then
           mForeColorNormal = RGB(5, 50, 89)
           mForeColorCheck = mForeColorNormal
           mForeColorDown = mForeColorNormal
           mForeColorUp = mForeColorNormal
           mForeColorDisabled = RGB(213, 216, 225)
           mBackColorSkinDefault = RGB(241, 243, 248)
           mBorderColorSkinDefault = RGB(178, 182, 197)
           Exit Sub
        End If
        
        
    End If
        
End Sub

' hDCSkin -> HDC en memoria con la imagen del Skin ( se usa en la función DrawSkin para dibujar la imagen de los diferentes estados del botón  )
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mSelectSkin(Optional ByVal lHandle As Long)
    If hDCSkin Then Call DeleteDC(hDCSkin) ' Eliminar el DC
    hDCSkin = CreateCompatibleDC(0)        ' Crearlo de nuevo
    Call SelectObject(hDCSkin, lHandle)    ' Establecer la imagen
End Sub



' Establecer BackColor
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mDrawBackColor(lcolor As Long)
    
    UserControl.BackColor = lcolor
    
    Exit Sub
    
    ' Creo que parpadea un poco menos
    
    'With UserControl
    '    If lHdc = 0 Then lHdc = .hdc
    '    If lWidth = 0 Then lWidth = .ScaleWidth
    '    If lHeight = 0 Then lHeight = .ScaleHeight
    'End With
    
    ' restángulo de dibujo ( Ancho y alto del botón )
    'Dim r As RECT
    'With r
    '    .Right = lWidth
    '    .Bottom = lHeight
    'End With
    
    'Dim hBrush As Long
    ' Crea el pincel del color indicado
    'hBrush = CreateSolidBrush(lColor)
    ' llenar el rectángulo con el pincel anterior
    'Call FillRect(lHdc, r, hBrush)
    ' liberar
    'Call DeleteObject(hBrush)
    
End Sub


' Función  para verificar los estados del mouse sobre el botón, y  redibujar
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mUpdateBtn(ByVal lMouseEvent As eMouseEvent)
    
    ' Flag Para que no redibuje en la carga de propiedades ( se  setea el flag al finalizar el ReadProperty para actualizae todo)
    If mFlagReadInitProp Or bFlagNoUpdateBtn Then Exit Sub
            
    ' guardar último estado, se usa para el refresh del UC y para el GotFocus
    lastStat = lMouseEvent
    
    ' guardar valores de la propiedad UnderLine por si se modifican mediante las propiedad useUnderLineMouseUp y para el UseUnderlineMousecheck
    Dim oldUnderLine As Boolean
    oldUnderLine = Font.UnderLine
    
    ' limpiar el UC
    UserControl.Cls
    
    ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Botón Deshabilitado ::: Enabled = False
    ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case mEnabled
       
       Case False
          
          If (mButtonType <> eNormalButton) And mValue Then
             If mUseUnderLineMouseCheck Then UserControl.FontUnderline = True
             Call DrawSkin(UserControl.hdc, , , TS_DISABLED_CHECKED)
          Else
              Call DrawSkin(UserControl.hdc, , , TS_DISABLED)
          End If
          Call mDrawCaption(Normal)
    
    ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Botón Habilitado ::: Enabled = True
    ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''
       Case True
          ' Estado normal, MouseUp, y MouseOver
          ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Select Case lMouseEvent
              Case 0
                If mbFlagMouseOver Then
                   Call DrawSkin(UserControl.hdc, , , TS_NORMAL)
                   If mValue And (mButtonType <> eNormalButton) Then
                      UserControl.FontUnderline = mUseUnderLineMouseCheck
                      Call DrawSkin(UserControl.hdc, , , TS_HOTCHECKED)
                      Call mDrawCaption(Abs(mbFlagMouseOver))
                   Else
                      Call DrawSkin(UserControl.hdc, , , TS_HOT)
                      UserControl.FontUnderline = mUseUnderLineMouseUp
                      Call mDrawCaption(Abs(mbFlagMouseOver))
                   End If
                   
                Else
                   ' MouseOut
                   ' ''''''''''''''''''''''''''''''''''''
                   Call DrawSkin(UserControl.hdc, , , TS_NORMAL)
                   If mValue And (mButtonType <> eNormalButton) Then
                      UserControl.FontUnderline = mUseUnderLineMouseCheck
                      Call DrawSkin(UserControl.hdc, , , TS_CHECKED)
                   End If
                   ' dibujar el texto del botón
                   Call mDrawCaption(Abs(mbFlagMouseOver))
                End If
          ' Estado MouseDown
          ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
              Case 2
               Call DrawSkin(UserControl.hdc, , , TS_PRESSED)
               UserControl.FontUnderline = mUseUnderLineMouseCheck
               Call mDrawCaption(Down)
          End Select
    End Select
    
    ' Restaurar El UnderLine ( Para las propiedades useUnderLineMouseUp y useUnderLineCheck )
    UserControl.FontUnderline = oldUnderLine
    
    If (lastStat = 0) And _
       mEnabled And _
       mbFocus Then
       
       Call mDrawFocus
    
    End If
    
End Sub




' Rutina para dibujar el texto del botón ( Propiedad Caption )
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mDrawCaption(ByVal lMouseEvent As eMouseEvent)
    
    ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Dibujar Caption con formato
    ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If mEnabledFormatText Then
       ' si se está en tiempo de diseño no es necesario redibujar el texto
       If Ambient.UserMode = False Then
          Exit Sub
       End If
       
       ' Si se presiona el botón, correr la posición X e Y del texto ( Variables mMargin y mTop de la posición para el texto con formato)
       If lMouseEvent = Down Then
          Dim oldMargin As Integer
          Dim oldTop    As Integer
          oldMargin = mMargin
          oldTop = mTop
          mMargin = mMargin + 3
          mTop = mTop + 1
          
          Call FormatTextDraw
          mTop = oldTop
          mMargin = oldMargin
       Else
          Call FormatTextDraw ' otros  eventos de mouse
       End If
       ' salir
       Exit Sub
    End If
    
    ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Dibujar la propiedad Caption
    ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim lcolor As Long
    
    
    ' colores de fuentes ....
    ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If mEnabled = False Then
       ' Color de fuente botón deshabilitado
       lcolor = mForeColorDisabled
    Else
       Select Case lMouseEvent
          ' normal
          Case 0
             If (mButtonType <> eNormalButton) And mValue Then
                lcolor = mForeColorCheck
             Else
                lcolor = mForeColorNormal
             End If
          ' MoueUp
          Case 1
             If (mButtonType <> eNormalButton) And mValue Then
                lcolor = mForeColorCheck
             Else
                lcolor = mForeColorUp
             End If
          ' MouseDown
          Case 2
             lcolor = mForeColorDown
       End Select
    End If
           
    With UserControl
        .ForeColor = lcolor
        ' si el Skin es de tipo enlace, usar el color de fondo del contenedor del botón
        If (mSkin = Link) And (mSkinEnabled) Then
           .BackColor = Extender.Container.BackColor
        End If
    End With
            
    ' Calcular área para dibujar el caption
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim x1 As Long
    Dim x2 As Long
    Dim y1 As Long
    Dim y2 As Long
    
    With UserControl
        Select Case mCaptionAlign
            Case 0: x1 = (.ScaleWidth - .TextWidth(mCaption)) / 2
            Case 1: x1 = mMarginCaption
            Case 2: x1 = (.ScaleWidth - .TextWidth(mCaption)) - mMarginCaption
        End Select
        If lMouseEvent = Down Then x1 = x1 + 1
        x2 = x1 + .TextWidth(mCaption)
        y1 = (.ScaleHeight - .TextHeight(mCaption)) / 2
        If lMouseEvent = Down Then y1 = y1 + 1
        y2 = y1 + .TextHeight(mCaption)
    End With
    
    ' Dibujar Textura - Verificar que hay imagen de textura, que hay caption, y que se está en estado normal
    If ((Not mPictureTexture Is Nothing) And _
       (lMouseEvent = Normal)) And _
       ((mCaption <> vbNullString)) Then
       
       Call mCreateTexture(x1, y1)
    ' Dibujar el Caption normal
    Else
       Dim r As RECT
       Call SetRect(r, x1, y1, x2, y2)  ' Copiar en r las coordenadas
        ' dibujar el texto
       Call DrawText(UserControl.hdc, mCaption, Len(mCaption), r, DT_CENTER)
       UserControl.Refresh  ' refrescar control
    End If
    
End Sub

' Función para dibujar un rectángulo ( para los bordes )
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mDrawRectangle(lcolor As Long)
    ' Si está habilitado el skin .. NO dibujar el borde
    If Not mSkinEnabled Then
        With UserControl
            .ForeColor = lcolor
            Call Rectangle(.hdc, 0, 0, .ScaleWidth, .ScaleHeight)
        End With
    End If
    
End Sub


' Función para establecer textura al caption del botón
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub mCreateTexture(lPosx As Long, lPosy As Long)

   Dim hDCMemory As Long
                        
   Dim lWidth As Long
   Dim lHeight As Long
                                                
                                                
   ' ancho y alto de la imagen de textura
   lWidth = ScaleX(mPictureTexture.Width, vbHimetric, vbPixels)
   lHeight = ScaleY(mPictureTexture.Height, vbHimetric, vbPixels)
   
   ' Comprobaciones
   With UserControl
       
       If .TextWidth(mCaption) > lWidth Then
          MsgBox "El ancho de la imagen de textura es menor al ancho del texto del botón. La imagen debe ser igual o mas grande", vbExclamation
          Exit Sub
       End If
       If .TextHeight(mCaption) > lHeight Then
          MsgBox "El alto de la imagen de textura es menor al alto del texto del botón. La imagen debe ser igual o mas grande", vbExclamation
          Exit Sub
       End If
       
       If .FontSize < 12 Then .FontSize = 12
       If .FontName = "MS Sans Serif" Then .FontName = "verdana"
   End With
   
   ' crear un DC temporal
   hDCMemory = CreateCompatibleDC(0)
                                        
   ' Seleccionar la imagen de textura en el DC anterior
   Call SelectObject(hDCMemory, mPictureTexture.Handle)
       
   With UserControl
      Call BeginPath(.hdc)
      Call TextOut(.hdc, lPosx, lPosy, mCaption, Len(mCaption)) ' dibujar texto en el UC
      Call EndPath(.hdc)
      Call SelectClipPath(.hdc, RGN_COPY) ' combinar los DC ( http://winapi.conclase.net/curso/index.php?fun=SelectClipPath)
      ' Copiar el hDCMemory --> Usercontrol.HDC
      Call BitBlt(.hdc, lPosx, lPosy, lWidth, lHeight, hDCMemory, 0, 0, &HCC0020)
      Call DeleteDC(hDCMemory) ' liberar recursos
      .Refresh ' refrescar el UC
   End With
End Sub


' Sub que calcula los valores de la fuente ( El alto y otros valores )
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mFormatTextCalcFont(ByRef pArrObject As tObjects)
    
    Dim ImgWidth            As Long
    Dim ImgHeight           As Long
    Dim tm                  As TEXTMETRIC
    Dim TextSize            As POINTAPI
            
    With UserControl
        ' Establecer la fuente
        Set .Font = pArrObject.oFont
    
        ' Recuperar los datos de la fuente
        Call GetTextMetrics(.hdc, tm)
        Call GetTextExtentPoint32(.hdc, pArrObject.sText, Len(pArrObject.sText), TextSize)
    End With
    
    With DataDrawText
       .lMaxAscentFont = tm.tmAscent
       .lMaxHeightFont = TextSize.Y
       .lWidthText = TextSize.X
       .lDesent = tm.tmDescent
    End With
    
    If DataDrawText.lDesent < tm.tmDescent Then
       DataDrawText.lDesent = tm.tmDescent
    End If
            
    If pArrObject.TypeObject = ePicture Then
       If Not pArrObject.Picture Is Nothing Then
          ImgWidth = ScaleX(pArrObject.Picture.Width, vbHimetric, vbPixels)
          ImgHeight = ScaleY(pArrObject.Picture.Height, vbHimetric, vbPixels)
          DataDrawText.lWidthText = ImgWidth
             
          If DataDrawText.lMaxAscentFont < ImgHeight Then
             DataDrawText.lMaxAscentFont = ImgHeight
          End If
          If DataDrawText.lMaxHeightFont < ImgHeight Then
             DataDrawText.lMaxHeightFont = ImgHeight + DataDrawText.lDesent
          End If
       End If
    End If
        
Exit Sub
error_Sub:
MsgBox Err.Description
End Sub


' Sub que corta las palabras del texto que se va a agregar y las agrega en el array arrObjects
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mFormatTextAdd(sText As String)
        
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim arrSplitText() As String
    
    sText = Replace(sText, vbTab, "   ")
    
    ' recuperar en un array las líneas
    arrSplitText = Split(sText, vbNewLine)
    ' recorrer el vector arrSplitText
    For j = LBound(arrSplitText) To UBound(arrSplitText)
        
        ' obtener las palabras
        Dim arrSplitWords() As String
        arrSplitWords = Split(arrSplitText(j), " ")
        ' recorrer las palabras
        For k = LBound(arrSplitWords) To UBound(arrSplitWords)
            If k < UBound(arrSplitWords) Then
                Call mFormatTextSetObject(arrSplitWords(k) & " ", eText) ' si no es la última agregarle un espacio
            Else
                Call mFormatTextSetObject(arrSplitWords(k), eText)
            End If
        Next
        
        If j < UBound(arrSplitText) Then
           Call mFormatTextSetObject("", eNewLine)
        End If
    Next
        
Exit Sub
error_Sub:

MsgBox Err.Description, vbCritical

End Sub


' Sub que agrega al array los elementos ( el texto, las imagenes, las lineas ) y las propieades
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mFormatTextSetObject(sValue As Variant, op As eObject)
    
    If (sValue = "") And op = eText Then Exit Sub
    
    ' redimensionar el array para el nuevo elemento
    Dim k As Long
    k = UBound(arrObjects) + 1
    ReDim Preserve arrObjects(k)
        
    With arrObjects(k)
        
        ' valores por defecto para que no de error al asignar las propiedades de la fuente
        If UserControl.Font.Size = 0 Then .oFont.Size = mDefFontSizeFormatText
        If UserControl.Font.Name = "" Then .oFont.Name = mDefFontNameFormatText
        
        Select Case op
            Case eNewLine ' Agregar nuevo salto de linea
                .sText = ""
                .bSaltoDeLinea = True
                .oFont.Size = mDefFontSizeFormatText
                .oFont.Name = mDefFontNameFormatText
                .TypeObject = eText
            Case eLine    ' Agregar los datos para luego dibujar una linea con el método Line
                .oFont.Size = mDefFontSizeFormatText
                .oFont.Name = mDefFontNameFormatText
                .lForeColor = sValue
                .TypeObject = eLine
                .bSaltoDeLinea = True
                .bLine.lDrawStyle = mLine.lDrawStyle
                .bLine.lDrawWidth = mLine.lDrawWidth
            Case ePicture ' Agregar los datos para esta imagen
                .sText = " "
                .TypeObject = ePicture
                .oFont.Size = mDefFontSizeFormatText
                .oFont.Name = mDefFontNameFormatText
                 Set .Picture = LoadPicture(sValue)
                .Align = mAlignText
            Case eText ' Agregar los datos para este texto
                Set .oFont = UserControl.Font
                .lBackColor = mBackColorFormatText
                .lForeColor = mForeColor
                .sText = sValue
                .Align = mAlignText
                .TypeObject = eText
        End Select
    End With
End Sub


' Función para averiguar cuando se produce el mouseOver y MouseOut en el botón ( Mientras se está en MouseOver, se inicializa el Timer para comprobar, y cuando se produce el mouseOut se detiene el Timer)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function mIsMouseOver() As Boolean
    Dim pt As POINTAPI
    ' coordendas del mouse
    Call GetCursorPos(pt)
    
    ' Retornar el Hwnd del UC si se está en mouseOver, si no devuelve False
    mIsMouseOver = (WindowFromPoint(pt.X, pt.Y) = hwnd)
End Function


' Cargar en el Array  las imágenes de Skin desde el archivo de recursos con el método LoadResPicture
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mLoadArrSkins(Optional lSkin As Long = -1)
    On Error Resume Next
    
    Dim lIndex As Integer
    lIndex = UBound(mArrstdPicSkins)
    
    If lSkin = -1 Then
        ' Si se produce error es por que no se habia cargado ... entonces cargar los gráficos
        If Err.Number Then
            ReDim mArrstdPicSkins(COUNT_SKIN_RES)
            Dim i As Integer
            For i = 0 To UBound(mArrstdPicSkins)
                Set mArrstdPicSkins(i) = LoadResPicture(101 + i, 0)
                'SavePicture mArrstdPicSkins(i), "c:\" & i & ".bmp"
            Next
        End If
    Else
        Erase mArrstdPicSkins
    End If
    On Error GoTo 0
    
End Sub



' Función para dibujar el rectángulo en el botón cuando tiene el foco ( Cuando se usan las teclas : Tab, teclas de dirección, o se llama a setFocus en tiempo de ejecución  )
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub mDrawFocus()

    If mShowFocusRect = False Then Exit Sub
    
    Dim r As RECT
    ' Llnenar en R, las dimensiones del UC
    With r
       .Left = 0
       .Top = 0
       .Bottom = ScaleHeight
       .Right = ScaleWidth
    End With
    
    ' dibujar el rectángulo
    With UserControl
        Call DrawFocusRect(.hdc, r)
        .Refresh
    End With
End Sub


' Timer que comprueba el MouseOut y MouseOver, y actualiza el estado del botón
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Timer1_Timer()
    ' si se produce el MouseOut .....
    If Not mIsMouseOver Then
       ' Lanzar evento, Terminar el timer y redibujar
       RaiseEvent MouseOut
       Timer1.Enabled = False
       mbFlagMouseOver = False
       Call mUpdateBtn(Normal)
    End If
End Sub




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'===============================================================================================================
' Fin de Funciones y Subs Privadas
'===============================================================================================================
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'===============================================================================================================
' Eventos UC
'===============================================================================================================
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' Evento del control de usuario cuando cambian propiedades varias del entorno ( para actualizar el BackColor cuando se usa mUseBackColorContainer = true , y para el skin de tipo Link)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If ((PropertyName = "BackColor") And (mUseBackColorContainer)) Then
       Call mUpdateBtn(Normal)
       
    ElseIf ((PropertyName = "BackColor") And (mSkin = Link)) Then
       UserControl.BackColor = Extender.Container.BackColor
       Call mDrawCaption(Normal)
    End If
End Sub



' Cuando recibe el foco redibujar en el botón el rectágnulo de enfoque
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserControl_GotFocus()
    Call mUpdateBtn(lastStat)
    If lastStat = 0 Then Call mDrawFocus
    mbFocus = True
End Sub


' Cuando pierde el foco
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserControl_LostFocus()
    mbFocus = False
    Call mUpdateBtn(lastStat)
End Sub


' Presión de teclas ( sacado del UC Chamaleon Button )
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
 
    RaiseEvent KeyDown(KeyCode, Shift)

    LastKeyDown = KeyCode
    Select Case KeyCode
    Case 32 'spacebar pressed
        Call UserControl_Click
    Case 39, 40 'right and down arrows
        SendKeys "{Tab}"
    Case 37, 38 'left and up arrows
        SendKeys "+{Tab}"
    End Select
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
    If (KeyCode = 32) And (LastKeyDown = 32) Then 'spacebar pressed, and not cancelled by the user
        If (mButtonType = eCheckbox) Then mValue = Not mValue
        Call mUpdateBtn(Normal)
        RaiseEvent Click
    End If
End Sub


' Inicializar variables por default
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserControl_Initialize()
    ' Cargar  imágenes de Skin desde el archivo de recursos
    Call mLoadArrSkins
    ' Valores para el texto con formato cuando no se indica FormatTextSetup antes de dibujarlo
    mTop = 3
    mMargin = 6
    mSpacingLine = 2
    
    ' Inicializar array con los datos del texto con formato
    ReDim arrObjects(0)
    
End Sub


Private Sub UserControl_Terminate()
    Call DeleteDC(hDCSkin)
    Call mLoadArrSkins(0)
    Set mSkinCustomPicture = Nothing
End Sub

' Escribir propiedades del control en el PropBag
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        
        Call .WriteProperty("PictureTexture", mPictureTexture, Nothing)
        Call .WriteProperty("ShowFocusRect", mShowFocusRect, False)
        Call .WriteProperty("UseBackColorContainer", mUseBackColorContainer, False)
        Call .WriteProperty("SkinCustomPicture", mSkinCustomPicture, Nothing)
        Call .WriteProperty("Skin", mSkin)
        Call .WriteProperty("Caption", mCaption)
        Call .WriteProperty("Enabled", mEnabled)
        Call .WriteProperty("Font", UserControl.Font)
        Call .WriteProperty("Value", mValue)
        
        Call .WriteProperty("ForeColorNormal", mForeColorNormal, mDefForeColorNormal)
        Call .WriteProperty("ForeColorDown", mForeColorDown, mDefForeColorDown)
        Call .WriteProperty("ForeColorUp", mForeColorUp, mDefForeColorUp)
        Call .WriteProperty("ForeColorDisabled", mForeColorDisabled, mDefForeColorDisabled)
        Call .WriteProperty("ForeColorCheck", mForeColorCheck, mDefForeColorCheck)
        
        Call .WriteProperty("CaptionAlign", mCaptionAlign)
        Call .WriteProperty("CaptionMargin", mMarginCaption)
        Call .WriteProperty("ButtonType", mButtonType)
        Call .WriteProperty("ToolTipText", mToolTipText)
        Call .WriteProperty("UseUnderLineMouseUp", mUseUnderLineMouseUp)
        Call .WriteProperty("UseUnderLineMouseCheck", mUseUnderLineMouseCheck)
        Call .WriteProperty("EnabledFormatText", mEnabledFormatText, False)
        Call .WriteProperty("BackColor", mBackColor, mDefBackColor)
        Call .WriteProperty("BackColorOver", mBackColorOver, mDefBackColorOver)
        Call .WriteProperty("BackColorDown", mBackColorDown, mDefBackColorDown)
        Call .WriteProperty("BackColorCheck", mBackColorCheck, mDefBackColorCheck)
        Call .WriteProperty("BackColorDisabled", mBackColorDisabled, mDefBackColorDisabled)
        Call .WriteProperty("BorderColorNormal", mBorderColorNormal, mDefBorderColorNormal)
        Call .WriteProperty("BorderColorOver", mBorderColorOver, mDefBorderColorOver)
        Call .WriteProperty("BorderColorDown", mBorderColorDown, mDefBorderColorDown)
        Call .WriteProperty("BorderColorCheck", mBorderColorCheck, mDefBorderColorCheck)
        Call .WriteProperty("BorderColorDisabled", mBorderColorDisabled, mDefBorderColorDisabled)
        Call .WriteProperty("SkinEnabled", mSkinEnabled, True)
        Call .WriteProperty("ColorSchemas", mColorSchemas, 0)
    End With


End Sub


' Leer las propiedades guardadas desde el PropBag
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    
    With PropBag
    
        ' flag para que no se produzca el mUpdateBtn mientras se leen ( Si se asignan a las variables locales igual no se asigna a la propiedad y no se produce el mUpdateBtn)
        mFlagReadInitProp = True
        
        Set PictureTexture = .ReadProperty("PictureTexture", Nothing)
        mUseBackColorContainer = .ReadProperty("UseBackColorContainer", False)
        mShowFocusRect = .ReadProperty("ShowFocusRect", False)
        mCaption = .ReadProperty("Caption", "")
        mEnabled = .ReadProperty("Enabled", True)
        mValue = .ReadProperty("Value", False)
        
        mForeColorNormal = .ReadProperty("ForeColorNormal", mDefForeColorNormal)
        mForeColorDown = .ReadProperty("ForeColorDown", mDefForeColorDown)
        mForeColorUp = .ReadProperty("ForeColorUp", mDefForeColorUp)
        mForeColorDisabled = .ReadProperty("ForeColorDisabled", mDefForeColorDisabled)
        mForeColorCheck = .ReadProperty("ForeColorCheck", mDefForeColorCheck)
        
        mCaptionAlign = .ReadProperty("CaptionAlign", 0)
        mMarginCaption = .ReadProperty("CaptionMargin", 0)
        mButtonType = .ReadProperty("ButtonType", 0)
        mToolTipText = .ReadProperty("ToolTipText", "")
        mUseUnderLineMouseUp = .ReadProperty("UseUnderLineMouseUp", False)
        mUseUnderLineMouseCheck = .ReadProperty("UseUnderLineMouseCheck", False)
        mEnabledFormatText = .ReadProperty("EnabledFormatText", False)
        Set UserControl.Font = .ReadProperty("Font", UserControl.Font)
        mBackColor = .ReadProperty("BackColor", mDefBackColor)
        mBackColorOver = .ReadProperty("BackColorOver", mDefBackColorOver)
        mBackColorDown = .ReadProperty("BackColorDown", mDefBackColorDown)
        mBackColorCheck = .ReadProperty("BackColorCheck", mDefBackColorCheck)
        mBackColorDisabled = .ReadProperty("BackColorDisabled", mDefBackColorDisabled)
        mBorderColorNormal = .ReadProperty("BorderColorNormal", mDefBorderColorNormal)
        mBorderColorOver = .ReadProperty("BorderColorOver", mDefBorderColorOver)
        mBorderColorDown = .ReadProperty("BorderColorDown", mDefBorderColorDown)
        mBorderColorCheck = .ReadProperty("BorderColorCheck", mDefBorderColorCheck)
        mBorderColorDisabled = .ReadProperty("BorderColorDisabled", mDefBorderColorDisabled)
        SkinEnabled = .ReadProperty("SkinEnabled", True)
        Set SkinCustomPicture = .ReadProperty("SkinCustomPicture", Nothing)
        
        ' Forzar la propiedad para que también actualice el redibujado del botón
        Skin = .ReadProperty("Skin", 0)
        mColorSchemas = .ReadProperty("ColorSchemas", 0)
        
    End With
    
    UserControl.Enabled = mEnabled
    
    ' Setear Flag para que ahora si se produzca el mUpdateBtn
    mFlagReadInitProp = False
    
    ' caargar valor de colores ( BackColor, color de fuentes, bordes, colores de fuente de cada Skin)
    Call mSetColorSchemas(mColorSchemas)
    ' Actualizar
    Call mUpdateBtn(Normal)
End Sub


' Evento cuando se redimensiona el botón ( Solo ejecutar el Resize en tiempo de diseño cuando se cambia las dimensiones o en tiempo de ejecución cuando se modifica las dimensiones des código)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserControl_Resize()
    If (Ambient.UserMode And IsWindowVisible(UserControl.hwnd) > 0) Or _
       (Ambient.UserMode = False) Then
        If mFlagReadInitProp = False Then
            Call mUpdateBtn(Normal)
        End If
    End If
End Sub



' Evento que solo se produce cuando se añade un nuevo control desde la ventana de controles de vb
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserControl_InitProperties()
           
   mFlagReadInitProp = True
   ShowFocusRect = False
   ForeColorNormal = vbBlack
   ForeColorUp = vbBlue
   ForeColorDown = vbRed
   ForeColorCheck = vbRed
   ForeColorDisabled = RGB(190, 190, 190)
   
   BackColor = Extender.Container.BackColor
   BackColorDown = Extender.Container.BackColor
   BackColorOver = Extender.Container.BackColor
   BackColorCheck = Extender.Container.BackColor
   BackColorDisabled = Extender.Container.BackColor
   
                               
   ' Setear propiedades por defecto
   ' '''''''''''''''''''''''''''''''''''''''
   
   ' Habilitar botón
   Enabled = True
   
   ' Cargar Los gráficos de Skins en el array
   Call mLoadArrSkins
   
   ' Habilitar el skin
   SkinEnabled = True
   
   ' skin por defecto
   Skin = Office2007BlueCheckGreen
   
   ' por defecto usar el color de esquemas de Skin
   ColorSchemas = useSkins
   
   ' No usar subrayados de fuente para los eventos de mouse
   UseUnderLineMouseUp = False
   UseUnderLineMouseCheck = False
   
   
   ' asignar al caption, el nombre del UC
   Caption = Extender.Name
   
   ' Margen en pixeles para el Caption
   CaptionMargin = 10
   
   ' Establecer en el Font la fuente del UC
   Set Font = UserControl.Font
   
   mFlagReadInitProp = False
   Call mUpdateBtn(Normal)
End Sub

' Si se está en tiempo de ejecución, redibujar el control
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserControl_Show()
    If Ambient.UserMode Then
       Call mUpdateBtn(Normal)
    End If
End Sub


' Eventos de Mouse del UC - Ejempplo sacado el control Chamaleon button
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    LastButton = Button
    If Button <> 2 Then
        Call mUpdateBtn(Down)
    End If
End Sub

' MouseUP
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If Button <> 2 Then
        Call mUpdateBtn(Normal)
    End If
End Sub

' Evento Clic
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserControl_Click()
    ' Vefificar si el botón está habilitado, y el estado de botón es válido
    If (LastButton = 1 Or LastKeyDown = 32) And mEnabled Then
        
       Select Case mButtonType
           
           ' Control checkBox
           ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''
           Case 1                                       '
                mValue = Not mValue
           ' Botón de opción
           ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''
           Case 2
                
                ' Referencia al formluario actual del UC
                Dim frm As Object
                Set frm = Extender.Parent
                
                ' Handle del contenedor ( Frame, picture, Form, otros ...)
                Dim lHwnd As Long
                lHwnd = Extender.Container.hwnd
                
                ' Recorrer todos los controles del formulario
                Dim ctrl As Control
                For Each ctrl In frm.Controls
                    With ctrl
                       ' verificar que sea el UC
                       If TypeOf ctrl Is ucBtnSkin Then
                          ' verificar que el ucBtnSkin es un OptionButton
                          If .ButtonType = 2 Then
                             ' si tiene el mismo Hwnd que este control, ponerlo en False para luego colocar en True el ucBtnSkin en el que se hizo clic
                             If (.Container.hwnd = lHwnd) And _
                                (ctrl.hwnd <> UserControl.hwnd) Then
                                
                                If .value Then .value = False
                                
                             End If
                          End If
                       End If
                    End With
                Next
                
                ' Si el optionButton estaba en False, setearlo a true
                If mValue = False Then mValue = True
                
       End Select
       
       ' Actualizar botón y lanzar evento Click
       Call mUpdateBtn(Normal)
       RaiseEvent Click
       
    End If
End Sub


'Evento MouseMove del UC
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo error_handler
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
    If Button < 2 Then
       If Not mIsMouseOver Then
          ' Esta línea  se ejecuta  Cuando se presiona el boton y sin soltarlo se saca el puntero del mouse fuera del botón
          Call mUpdateBtn(Normal)
       Else
          If Button = 0 And Not mbFlagMouseOver Then
             ' Este bloque  se ejecuta  cuando se entra al botón, y se activa el timer para saber cuando se produce el MouseOut
             Timer1.Enabled = True
             mbFlagMouseOver = True
             Call mUpdateBtn(Normal)
             RaiseEvent MouseOver
          ElseIf Button = 1 Then
             mbFlagMouseOver = True
             Call mUpdateBtn(Down)
             mbFlagMouseOver = False
          End If
       End If
    End If
    
error_handler:
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'===================================================================================================
' Fin de Eventos del User control
'===================================================================================================
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'///////////////////////////////////////////////////////////////////////////////////////////////////
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'===================================================================================================
' Propiedades
'===================================================================================================
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' Propiedad para establecer el estado normal del botón con el color que tenga el contenedor ( Para cuando tiene Skin o sin Skin)
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get UseBackColorContainer() As Boolean
    UseBackColorContainer = mUseBackColorContainer
End Property

Property Let UseBackColorContainer(bValue As Boolean)

    mUseBackColorContainer = bValue
    Call PropertyChanged("UseBackColorContainer")
    
    Call mUpdateBtn(Normal)
End Property



' Propiedad para la Fuente
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByRef newFont As Font)

    If Ambient.UserMode = False And mEnabledFormatText Then
       MsgBox "No se puede cambiar el valor de esta propiedad si la propiedad EnabledFormatText se encuentra en 'True' ", vbExclamation
    Else
       Set UserControl.Font = newFont
       Call mUpdateBtn(Normal)
       Call PropertyChanged("Font")
    End If
End Property




' Negrita
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_MemberFlags = "400"
    FontBold = UserControl.FontBold
End Property

Public Property Let FontBold(ByVal newValue As Boolean)
    UserControl.FontBold = newValue
    ' Si no se usa texto con formato , redibujar el control ( Fondo y caption )
    If mEnabledFormatText = False Then
       Call mUpdateBtn(Normal)
    End If
End Property


' Italica
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_MemberFlags = "400"
    FontItalic = UserControl.FontItalic
End Property

Public Property Let FontItalic(ByVal newValue As Boolean)
    UserControl.FontItalic = newValue
    ' Si no se usa texto con formato , redibujar el control ( Fondo y caption )
    If mEnabledFormatText = False Then
       Call mUpdateBtn(Normal)
    End If
End Property


' Subrayado
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_MemberFlags = "400"
    FontUnderline = UserControl.FontUnderline
End Property

Public Property Let FontUnderline(ByVal newValue As Boolean)
    UserControl.FontUnderline = newValue
    ' Si no se usa texto con formato , redibujar el control ( Fondo y caption )
    If mEnabledFormatText = False Then
       Call mUpdateBtn(Normal)
    End If
End Property


' Tamaño de fuente
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get FontSize() As Integer
Attribute FontSize.VB_MemberFlags = "400"
    FontSize = UserControl.FontSize
End Property

Public Property Let FontSize(ByVal newValue As Integer)
    UserControl.FontSize = newValue
    ' Si no se usa texto con formato , redibujar el control ( Fondo y caption )
    If mEnabledFormatText = False Then
       Call mUpdateBtn(Normal)
    End If
End Property

' Nombre de fuente
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get FontName() As String
Attribute FontName.VB_MemberFlags = "400"
    FontName = UserControl.FontName
End Property

Public Property Let FontName(ByVal newValue As String)
    
    On Error Resume Next
    With UserControl
        .FontName = newValue
        If Err.Number <> 0 Then
           .FontName = Ambient.Font.Name
        End If
    End With
    On Error GoTo 0
    
    ' Si no se usa texto con formato , redibujar el control ( Fondo y caption )
    If mEnabledFormatText = False Then
       Call mUpdateBtn(Normal)
    End If
    
End Property


' Texto del botón
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get Caption() As String
    Caption = mCaption
End Property

Public Property Let Caption(ByVal newValue As String)
    mCaption = newValue
    ' Si no se usa texto con formato , redibujar el control ( Fondo y caption )
    If Not mEnabledFormatText Then
       Call mUpdateBtn(Normal)
    End If
    Call PropertyChanged("Caption")
End Property


' Enabled
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    mEnabled = newValue
    UserControl.Enabled = mEnabled
    ' Setear colores ( bordes, fuente y colores de fondos )
    Call mSetColorSchemas(mColorSchemas)
    ' Redibujar
    Call mUpdateBtn(Normal)
    Call PropertyChanged("Enabled")
End Property


' Propiedad para habilitar o deshabilitar los skins
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get SkinEnabled() As Boolean
    SkinEnabled = mSkinEnabled
End Property

Property Let SkinEnabled(bValue As Boolean)

    mFlagSkinEnabled = True
    
    mSkinEnabled = bValue
    
    If bValue Then
       If mSkin = CustomSkin Then
          mColorSchemas = anone     ' Colores de fuente personalizados para cuando se usa un Skin propio
       Else
          mColorSchemas = useSkins  ' esquema de colores cuando no se  usan los personalizados
       End If
    Else
       mColorSchemas = anone     ' si se deshabilita el skin, establecer el esquema en Ninguno .. para poder definir los colores ( backcolor, forecolor y borde)
    End If

    Call PropertyChanged("SkinEnabled")
    Skin = mSkin                    ' Forzar el cambio de Skin para que actualice la propiedad ColorSchemas
    mFlagSkinEnabled = False        ' Flag para que no entre en un bucle infinito
    
    Call mUpdateBtn(Normal)
    
End Property


' Propiedad para los Skins
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get Skin() As eSkin
    Skin = mSkin
End Property

Public Property Let Skin(ByVal newValue As eSkin)
    

    
    ' Si el Skin está deshabilitado .. mostrar un mensaje cuando se está en tiempo de diseño. y salir de la propiedad
    If (mSkinEnabled = False) And _
       (mFlagSkinEnabled = False) And _
       (mFlagReadInitProp = False) Then
       If Not Ambient.UserMode Then
          MsgBox "La propiedad SkinEnabled se encuentra en False. Para cambiar un skin primero establezca la propiedad SkinEnabled en 'True' ", vbExclamation
       End If
       
       Exit Property
    
    End If
    
    If mEnabledFormatText And newValue = Link Then
       MsgBox "La opción de Link no se puede usar cuando la propiedad EnabledFormatText está en True ", vbExclamation
       Exit Property
    End If
    
    ' asignar nuevo skin
    mSkin = newValue
    
    ' .. si el skin está habilitado
    If (mSkinEnabled) And _
       (mSkin <> CustomSkin) And _
       (mSkin <> Link) Then
        ' establecer el esquema para los colores por defecto de cada skin que luego se cargan desde mSetColorSchemas
        mColorSchemas = useSkins
        ' Pasar la imagen ( ya está cargada en el array .. pero por las dudas)
        Call mSelectSkin(mArrstdPicSkins(mSkin).Handle)
    
    ElseIf (mSkin = CustomSkin) And _
           (Not mSkinCustomPicture Is Nothing) Then
        mColorSchemas = anone                       ' Cambiar esquema a colores personalizados cuando se usa un skin propio
        Call mSelectSkin(mSkinCustomPicture.Handle)
    ElseIf mSkin = Link Then
        mColorSchemas = anone
    End If
    
    ' Recuperar colores de fuente, backcolor y redibujar
    Call PropertyChanged("Skin")
    Call mSetColorSchemas(mColorSchemas)
    Call mUpdateBtn(Normal)
    
End Property


' Color de fuente
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get ForeColorNormal() As OLE_COLOR
    ForeColorNormal = mForeColorNormal
End Property

Property Let ForeColorNormal(lValue As OLE_COLOR)
    
    If Ambient.UserMode = False And mSkinEnabled And mColorSchemas = useSkins Then
       Call mShowErrorChangeProp
    Else
       mForeColorNormal = lValue
       Call PropertyChanged("ForeColorNormal")  'Guardar valor de propiedad en el PropBag
       ' Si no se usa texto con formato , redibujar el control ( Fondo y caption )
       If Not mEnabledFormatText Then
          Call mUpdateBtn(Normal)
       End If
    End If

End Property

' Propiedad para el color de fuente cuando se hace el MouseUP
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get ForeColorUp() As OLE_COLOR
    ForeColorUp = mForeColorUp
End Property
Property Let ForeColorUp(lValue As OLE_COLOR)
    If Ambient.UserMode = False And mColorSchemas = useSkins Then
       Call mShowErrorChangeProp
    Else
        mForeColorUp = lValue
        Call PropertyChanged("ForeColorUp")
        ' Si no se usa texto con formato , redibujar el control ( Fondo y caption )
        If mEnabledFormatText = False Then
           Call mUpdateBtn(Normal)
        End If
    End If
End Property


' Propiedad para el color de fuente cuando se hace el MouseDown
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get ForeColorDown() As OLE_COLOR
    ForeColorDown = mForeColorDown
End Property

Property Let ForeColorDown(lValue As OLE_COLOR)
    If Ambient.UserMode = False And mColorSchemas = useSkins Then
       Call mShowErrorChangeProp
    Else
        mForeColorDown = lValue
        Call PropertyChanged("ForeColorDown")
        ' Si no se usa texto con formato , redibujar el control ( Fondo y caption )
        If mEnabledFormatText = False Then
           Call mUpdateBtn(Normal)
        End If
    End If
End Property


' Propiedad para el color de fuente cuando el botón está deshabilitado
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get ForeColorDisabled() As OLE_COLOR
    ForeColorDisabled = mForeColorDisabled
End Property

Property Let ForeColorDisabled(lValue As OLE_COLOR)
    
    If Ambient.UserMode = False And mColorSchemas = useSkins Then
       Call mShowErrorChangeProp
    Else
        mForeColorDisabled = lValue
        PropertyChanged "ForeColorDisabled"
        If mEnabledFormatText = False Then
           Call mUpdateBtn(Normal)
        End If
    End If
End Property


' Propiedad para el color de fuente cuando el botón es un OptionButton / Check y se encuentra el value en True
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get ForeColorCheck() As OLE_COLOR
    ForeColorCheck = mForeColorCheck
End Property

Property Let ForeColorCheck(lValue As OLE_COLOR)
    If Ambient.UserMode = False And mColorSchemas = useSkins Then
       Call mShowErrorChangeProp
    Else
       mForeColorCheck = lValue
       Call PropertyChanged("ForeColorCheck")
       If mEnabledFormatText = False Then
          Call mUpdateBtn(Normal)
       End If
    End If
End Property


' Propiedad Value para cuando el botón es un OptionButton o un checkBox
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get value() As Boolean
    value = mValue
End Property

Public Property Let value(ByVal newValue As Boolean)

    mValue = newValue
    ' verificar el estilo de botón , y si no es normal redibujar
    If (mButtonType = eCheckbox Or mButtonType = eOptionbutton) Then
       Call mUpdateBtn(Normal)
       If Ambient.UserMode Then
          ' si se está en tiempo de ejecución lanzar el evento click al modificar el value del control
          RaiseEvent Click
       End If
    End If
    Call PropertyChanged("Value")
End Property


' Alineación del texto del botón ( Izquierda, derecha y Centro )
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get CaptionAlign() As eCaptionAlignment
    CaptionAlign = mCaptionAlign
End Property

Property Let CaptionAlign(lValue As eCaptionAlignment)
    mCaptionAlign = lValue
    ' Si no se usa texto con formato , redibujar el control ( Fondo y caption )
    If mEnabledFormatText = False Then
       Call mUpdateBtn(Normal)
    End If
    PropertyChanged "CaptionAlign"
End Property



' Margen en pixeles para el texto del botón ( Texto normal, no con formato )
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get CaptionMargin() As Long
    CaptionMargin = mMarginCaption
End Property

Property Let CaptionMargin(lValue As Long)
    If lValue <= 1 Then lValue = 1              ' setear el mínimo valor
    mMarginCaption = lValue
    ' Si no se usa texto con formato , redibujar el control ( Fondo y caption )
    If mEnabledFormatText = False Then
       Call mUpdateBtn(Normal)
    End If
    Call PropertyChanged("CaptionMargin")
End Property


' Propiedad que indica el tipo de botón ( 0 - Botón normal, 1 - CheckBox - 2 OptionButton)
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get ButtonType() As eButtonType
    ButtonType = mButtonType
End Property

Property Let ButtonType(lValue As eButtonType)
    mButtonType = lValue
    ' Actualizar
    Call mUpdateBtn(Normal)
    Call PropertyChanged("ButtonType")
End Property


' Propiedad ToolTipText
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Let ToolTipText(sValue As String)
    mToolTipText = sValue
    Call PropertyChanged("ToolTipText")
End Property

Property Get ToolTipText() As String
    ToolTipText = mToolTipText
End Property


' Propiedad de solo lectura con el Handle del UC
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get hwnd() As Long
Attribute hwnd.VB_MemberFlags = "400"
    hwnd = UserControl.hwnd
End Property


' Propiedad para subrayar el texto del botón cuando se realiza un MouseUp
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get UseUnderLineMouseUp() As Boolean
     UseUnderLineMouseUp = mUseUnderLineMouseUp
End Property

Property Let UseUnderLineMouseUp(bValue As Boolean)
     
     ' No usar subrayado cuando se usa texto con formato
     If mEnabledFormatText And Ambient.UserMode = False Then
        MsgBox "No se puede cambiar el valor de esta propiedad si la propiedad EnabledFormatText se encuentra en 'True'  ", vbExclamation
     Else
        mUseUnderLineMouseUp = bValue
        Call PropertyChanged("UseUnderLineMouseUp")
     End If

End Property


' Propiedad para subrayar el texto del botón cuando es un Option Button o CheckBox y el valor de 'Value' es True, o se hace un MouseDown
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get UseUnderLineMouseCheck() As Boolean
     UseUnderLineMouseCheck = mUseUnderLineMouseCheck
End Property

Property Let UseUnderLineMouseCheck(bValue As Boolean)
     
     ' No usar subrayado cuando se usa texto con formato
     If mEnabledFormatText And Ambient.UserMode = False Then
        MsgBox "No se puede cambiar el valor de esta propiedad si la propiedad EnabledFormatText se encuentra en 'True'  ", vbExclamation
     Else
        mUseUnderLineMouseCheck = bValue
        Call PropertyChanged("UseUnderLineMouseCheck")
     
        Call mSetColorSchemas(mColorSchemas)
        Call mUpdateBtn(0)
     End If
     
End Property


' Propiedad de solo lectura con el Hdc del UC
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get hdc() As Long
    hdc = UserControl.hdc
End Property


'Propiedad para habilitar/deshabilitar el texto con formato para el botón, si está en False se usa el texto del valor de la propiedad Caption
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get EnabledFormatText() As Boolean
    EnabledFormatText = mEnabledFormatText
End Property
Property Let EnabledFormatText(bValue As Boolean)
    
    mEnabledFormatText = bValue
    
    ' verificar que no se está utilizando texto con formato para usar el subrayado del caption
    ' si el subrayado estaba activo, deshabilitarlo ...
    
    If mEnabledFormatText Then
       If mUseUnderLineMouseUp Then
          mUseUnderLineMouseUp = False
          Call PropertyChanged("UseUnderLineMouseUp")   ' Guardar el valor de la propiedad en el PropBag
       End If
       If mUseUnderLineMouseCheck Then
          mUseUnderLineMouseCheck = False
          Call PropertyChanged("UseUnderLineMouseCheck") ' Guardar el valor de la propiedad en el PropBag
       End If
    End If
    
    ' Resetear el grosor y el estilo del lapiz de los métodos gráficos del UC, cuando no se usa el texto con formato
    If bValue = False Then
       With UserControl
          .DrawWidth = 1
          .DrawStyle = 0
       End With
       Call mUpdateBtn(Normal)                 ' Cuando se Deshabilita , Actualizar el valor del 'Caption'
    Else
        ' Solo redibujar cuando se está en tiempo de ejecución
       If Ambient.UserMode Then
          Call mUpdateBtn(Normal)
          Call FormatTextDraw
       End If
    End If
    
    ' Guardar el cambio de propiedad
    Call PropertyChanged("EnabledFormatText")
    
End Property


' Propiedad para el color de fondo cuando no se usa un Skin
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal lValue As OLE_COLOR)
    
    If Ambient.UserMode = False And mSkin = Link Then
       Call mShowErrorChangeProp3
    ElseIf Ambient.UserMode = False And mColorSchemas = useSkins Then
       Call mShowErrorChangeProp
    ElseIf Ambient.UserMode = False And mSkinEnabled Then
       Call mShowErrorChangeProp2
    Else
       mBackColor = lValue
       Call PropertyChanged("BackColor")
       Call mUpdateBtn(0) ' Actualizar
    End If

End Property


' Propiedad para el color de fondo al hacer un MouseUp, y cuando no se usa un Skin
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get BackColorOver() As OLE_COLOR
    BackColorOver = mBackColorOver
End Property

Public Property Let BackColorOver(ByVal lValue As OLE_COLOR)
    
    If Ambient.UserMode = False And mSkin = Link Then
       Call mShowErrorChangeProp3
    ElseIf Ambient.UserMode = False And mColorSchemas = useSkins Then
       Call mShowErrorChangeProp
    ElseIf Ambient.UserMode = False And mSkinEnabled Then
       Call mShowErrorChangeProp2
    Else
       mBackColorOver = lValue
       Call PropertyChanged("BackColorOver")
    End If
    
End Property


' Propiedad para el color de fondo al hacer un MouseDown, y cuando no se usa un Skin
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get BackColorDown() As OLE_COLOR
    BackColorDown = mBackColorDown
End Property

Public Property Let BackColorDown(ByVal lValue As OLE_COLOR)
    
    If Ambient.UserMode = False And mSkin = Link Then
       Call mShowErrorChangeProp3
    ElseIf Ambient.UserMode = False And mColorSchemas = useSkins Then
       Call mShowErrorChangeProp
    ElseIf Ambient.UserMode = False And mSkinEnabled Then
       Call mShowErrorChangeProp2
    Else
       mBackColorDown = lValue
       Call PropertyChanged("BackColorDown")
    End If

End Property

' Propiedad para el color de fondo cuando se encuentra el botón chequeado, y cuando no se usa un Skin
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get BackColorCheck() As OLE_COLOR
    BackColorCheck = mBackColorCheck
End Property

Public Property Let BackColorCheck(ByVal lValue As OLE_COLOR)
    
    If Ambient.UserMode = False And mSkin = Link Then
       Call mShowErrorChangeProp3
    ElseIf Ambient.UserMode = False And mColorSchemas = useSkins Then
       Call mShowErrorChangeProp
    ElseIf Ambient.UserMode = False And mSkinEnabled Then
       Call mShowErrorChangeProp2
    Else
       mBackColorCheck = lValue
       Call PropertyChanged("BackColorCheck")
       Call mUpdateBtn(0)
    End If

End Property


' Propiedad para el color de fondo cuando se encuentra el botón Deshabilitado, y cuando no se usa un Skin
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get BackColorDisabled() As OLE_COLOR
    BackColorDisabled = mBackColorDisabled
End Property

Public Property Let BackColorDisabled(ByVal lValue As OLE_COLOR)
    
    If Ambient.UserMode = False And mSkin = Link Then
       Call mShowErrorChangeProp3
    ElseIf Ambient.UserMode = False And mColorSchemas = useSkins Then
       Call mShowErrorChangeProp
    ElseIf Ambient.UserMode = False And mSkinEnabled Then
       Call mShowErrorChangeProp2
    Else
        mBackColorDisabled = lValue
        Call PropertyChanged("BackColorDisabled")
        Call mUpdateBtn(0)
    End If

End Property

' Bordes
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Borde normal
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get BorderColorNormal() As OLE_COLOR
    BorderColorNormal = mBorderColorNormal
End Property

Public Property Let BorderColorNormal(ByVal lValue As OLE_COLOR)
    
    If Ambient.UserMode = False And mSkin = Link Then
       Call mShowErrorChangeProp3
    ElseIf Ambient.UserMode = False And mColorSchemas = useSkins Then
       Call mShowErrorChangeProp
    ElseIf Ambient.UserMode = False And mSkinEnabled Then
       Call mShowErrorChangeProp2
    Else
        mBorderColorNormal = lValue
        Call PropertyChanged("BorderColorNormal")
        Call mUpdateBtn(0)
    End If

End Property

' Color de Borde para cuando se está encima del botón
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get BorderColorOver() As OLE_COLOR
    BorderColorOver = mBorderColorOver
End Property
Public Property Let BorderColorOver(ByVal lValue As OLE_COLOR)
    
    If Ambient.UserMode = False And mSkin = Link Then
       Call mShowErrorChangeProp3
    ElseIf Ambient.UserMode = False And mColorSchemas = useSkins Then
       Call mShowErrorChangeProp
    ElseIf Ambient.UserMode = False And mSkinEnabled Then
       Call mShowErrorChangeProp2
    Else
       mBorderColorOver = lValue
       Call PropertyChanged("BorderColorOver")
       Call mDrawRectangle(lValue)
    End If

End Property


' Color de Borde para cuando se presiona el botón
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get BorderColorDown() As OLE_COLOR
    BorderColorDown = mBorderColorDown
End Property

Public Property Let BorderColorDown(ByVal lValue As OLE_COLOR)
    
    If Ambient.UserMode = False And mSkin = Link Then
       Call mShowErrorChangeProp3
    ElseIf Ambient.UserMode = False And mColorSchemas = useSkins Then
       Call mShowErrorChangeProp
    ElseIf Ambient.UserMode = False And mSkinEnabled Then
       Call mShowErrorChangeProp2
    Else
        mBorderColorDown = lValue
        Call PropertyChanged("BorderColorDown")
        Call mDrawRectangle(lValue) ' Dibujar rectángulo
    End If

End Property


' Color de Borde para cuando el Value está en true ( Para OptionButtons y CheckBox)
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get BorderColorCheck() As OLE_COLOR
    BorderColorCheck = mBorderColorCheck
End Property

Public Property Let BorderColorCheck(ByVal lValue As OLE_COLOR)
    
    If Ambient.UserMode = False And mSkin = Link Then
       Call mShowErrorChangeProp3
    ElseIf Ambient.UserMode = False And mColorSchemas = useSkins Then
       Call mShowErrorChangeProp
    ElseIf Ambient.UserMode = False And mSkinEnabled Then
       Call mShowErrorChangeProp2
    Else
        mBorderColorCheck = lValue
        Call PropertyChanged("BorderColorCheck")
    
        Call mDrawRectangle(lValue) ' Dibujar rectángulo
    End If

End Property


' Color de Borde para cuando el botón está deshabilitado
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get BorderColorDisabled() As OLE_COLOR
    BorderColorDisabled = mBorderColorDisabled
End Property

Public Property Let BorderColorDisabled(ByVal lValue As OLE_COLOR)
    
    If Ambient.UserMode = False And mSkin = Link Then
       Call mShowErrorChangeProp3
    ElseIf Ambient.UserMode = False And mColorSchemas = useSkins Then
       Call mShowErrorChangeProp
    ElseIf Ambient.UserMode = False And mSkinEnabled Then
       Call mShowErrorChangeProp2
    Else
       mBorderColorDisabled = lValue
       Call PropertyChanged("BorderColorDisabled")
       Call mUpdateBtn(0)
    End If

End Property


' Propiedad para guardar los esquemas de color ( Skins, o None para definir los que se quieran)
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get ColorSchemas() As eColorSchemas
    ColorSchemas = mColorSchemas
End Property

Property Let ColorSchemas(lValue As eColorSchemas)
    
    ' si el Skin está habilitado, no cambiar el esquema de colores
    If ((lValue <> useSkins) And mSkinEnabled) And _
       (lValue <> anone) Then
       MsgBox "Esta opción es para usar cuando la propiedad SkinEnabled está desactivada. Para usar colores personalizados de fuente, borde y fondo para botones sin Skin, seleccionae la opción [None] de la propiedad ColorSchema, o establezca la opción SkinEnabled en False", vbInformation
       Exit Property ' salir
    End If
    
    mColorSchemas = lValue
    
    'Actualizar el esquema de colores y redibujar el botón
    Call PropertyChanged("ColorSchemas")
    Call mSetColorSchemas(mColorSchemas)
    Call mUpdateBtn(Normal)
    
End Property


' Propiedad para cargar un Skin propio, ya sea desde la ventana de propiedades o e tiempo de ejecución desde el disco, archivo de recursos etc ..
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get SkinCustomPicture() As Picture
    Set SkinCustomPicture = mSkinCustomPicture
End Property

Property Set SkinCustomPicture(ByVal stdPic As Picture)
    
    Set mSkinCustomPicture = stdPic
    
    If Not mSkinCustomPicture Is Nothing Then
        Me.bFlagNoUpdateBtn = True
        ColorSchemas = anone            ' Cambiar el esquema de colores, a None para poder seleccionar los colores que se quieran
        Me.bFlagNoUpdateBtn = False
        Skin = CustomSkin
    Else
        Call mSelectSkin(0)            ' Si mSkinCustomPicture no tiene una imagen, pasar como handle un 0 para borrar la imagen de hDCSkin
    End If
    
    ' Actualizar
    Call PropertyChanged("SkinCustomPicture")
    Call mSetColorSchemas(mColorSchemas)
    Call mUpdateBtn(Normal)
    
End Property


' Propiedad de tipo stdPicture para la textura del caption del botón
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get PictureTexture() As Picture
    Set PictureTexture = mPictureTexture
End Property

Property Set PictureTexture(stdPicValue As Picture)
    
    Set mPictureTexture = stdPicValue
    
    ' si está habilitado el texto con formato , no aplicar textura, pero si guardarla
    If Not mEnabledFormatText Then
       Call mUpdateBtn(0)
    End If
    ' notificar en el PropBag
    Call PropertyChanged("PictureTexture")
    
End Property

' ( BackColor y color de borde) - Propiedades para usar en tiempo de ejecución para asignarlos a otros controles ( BackColor del Form etc ..)
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Property Get BackColorSkinDefault() As Long
    BackColorSkinDefault = mBackColorSkinDefault
End Property

Property Get BorderColorSkinDefault() As Long
    BorderColorSkinDefault = mBorderColorSkinDefault
End Property

Property Get ShowFocusRect() As Boolean
    ShowFocusRect = mShowFocusRect
End Property

Property Let ShowFocusRect(bValue As Boolean)
    mShowFocusRect = bValue
    Call mUpdateBtn(0)
    Call PropertyChanged("ShowFocusRect")
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'===================================================================================================
' Fin de Propiedades
'===================================================================================================
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
