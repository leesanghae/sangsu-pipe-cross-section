Attribute VB_Name = "횡단면도그리기"
Option Explicit

'=====================================================================================
'  상수관로 횡단면도 자동작도   (AutoCAD / ZWCAD 공용)
'  ------------------------------------------------------------------------------
'  · 토적표(표 값 + 행 순서) + SDB(H)(기하) 를 읽어 A1 33장에 655단면 작도
'  · "포인트 선택 → 일률 작도" : 기준점 1회 클릭 후 전 단면 자동 작도
'  · 모든 변수는 'P' 시트에서 수정 (코드 수정 불필요)
'
'  진입점 (매크로 목록에서 실행)
'    횡단_P시트생성    : 파라미터 시트를 기본값으로 생성 (최초 1회)
'    횡단_1단계_검증   : 데이터 정합성만 검사 → '검토_검증' 시트 출력
'    횡단_작도         : 검증 → 포인트 선택 → 전체 작도
'    횡단_블록갱신     : 기존 관로06 속성만 갱신 (도면 재작성 없이)
'    횡단_도곽확인     : 도면의 블록/도곽(XREF) 상태 점검
'=====================================================================================

'────────────────────────── 0. 상수 ──────────────────────────
' [중요] 모듈 레벨 Const/Dim/Type/Enum 은 반드시 '선언부'(모든 프로시저 앞)에만 둔다.
'        프로시저 사이에 놓으면 VBA 가 인식하지 못해
'        "컴파일 오류: 변수가 정의되지 않았습니다" 가 발생한다. (v2.2 수정)
Private Const PSHEET As String = "P"
Private Const VSHEET As String = "검토_검증"
Private Const PI As Double = 3.14159265358979
Private Const TINY As Double = 0.0000001

' AL_* : 문자 정렬 (AcText.Alignment 값)  0=좌하  9=중앙좌  10=중앙중심  11=중앙우  13=하중앙
Private Const AL_LEFT As Long = 0
Private Const AL_MLFT As Long = 9
Private Const AL_MCTR As Long = 10
Private Const AL_MRGT As Long = 11
Private Const AL_BCTR As Long = 13

'────────────────────────── 1. 단면 데이터 ──────────────────────────
Private sN As Long
Private sSta() As String          ' 측점 문자열 (토적표 B열 그대로)
Private sGH() As Double           ' 지반고
Private sInv() As Double          ' 관저고
Private sDia() As Double          ' 관경 mm
Private sCode() As String         ' OP02 / OP04 / OP01 / SC01 ...
Private sSlope() As Double        ' 굴착경사 (SDB(H) G열)
' 표 블록 13개 값 (토적표)
Private sExc() As Double, sWet() As Double, sAr() As Double, sAb() As Double
Private sAsp() As Double, sCon() As Double, sAsc() As Double, sBodo() As Double
Private sSand() As Double, sSub() As Double, sFrost() As Double
' SDB(H) 원본 (매칭용)
Private hN As Long
Private hSta() As String, hGH() As Double, hInv() As Double, hDia() As Double
Private hCode() As String, hSlope() As Double
Private sB() As Double            ' 토적표 굴착폭 (CU)
' 매칭 결과
Private sSrc() As String          ' "SDBH" / "TOKJUK"
Private sVerIdx() As Long         ' 검증용 토적표 행번호
Private sXD() As Variant          ' XDATA 8값 (토적표 CP..CW)

' 단면 1개의 계산 결과  (UDT 대신 모듈 배열 - VBA 호환성 우선)
Private gD As Double, gH As Double, gb0 As Double, gtv As Double
Private gzL As Double, gzP As Double, gB As Double
Private ght As Double, ghb As Double
Private gExc As Double, gWet As Double, gAr As Double, gAb As Double
Private gSand As Double, gSub As Double, gFrost As Double
Private gAsp As Double, gCon As Double, gAsc As Double, gBodo As Double
Private gTopik As Double, gHasExc As Boolean
Private gzRfil As Double, gRfilHalf As Double
Private gNLay As Long
Private gLayName() As String, gLayTh() As Double
Private gLayKind() As String, gLayDisp() As String
' [v2.4] 아래 2개는 원래 SecCalc 앞(프로시저 뒤)에 선언되어 있어
'        Wz() 안의 gSlp 사용처에서 "변수가 정의되지 않았습니다" 가 났다.
Private gCurK As Long
Private gSlp As Double

'────────────────────────── 2. 파라미터 ──────────────────────────
' [배치]
Private pNCol As Long, pNRow As Long
Private pSheetW As Double, pPitchX As Double, pPitchY As Double
Private pOffX As Double, pOffY As Double
Private pPickBase As Boolean, pBaseX As Double, pBaseY As Double
' [소스]
Private pSrcGeom As String        ' SDBH / TOKJUK / AUTO
' [블록]
Private pBlkTable As String, pBlkBorder As String, pBlkTitle As String, pBlkDot As String
Private pBlkScale As Double, pDotScale As Double, pTblDX As Double
Private pBorderXref As String, pBorderW As Double, pBorderH As Double, pBorderFallback As Boolean
' [레이어]
Private pLayCent As String, pLayGrnd As String, pLayPave As String, pLayExcv As String
Private pLayRfil As String, pLayPipe As String, pLayPatt As String, pLayLead As String
Private pLayLeat As String, pLayLeal As String, pLayDiml As String, pLayText As String
Private pLayTabl As String, pLayBorder As String, pLayTitle As String
Private cCent As Long, cGrnd As Long, cPave As Long, cExcv As Long, cRfil As Long
Private cPipe As Long, cPatt As Long, cLead As Long, cLeat As Long, cLeal As Long
Private cDiml As Long, cText As Long, cTabl As Long, cBorder As Long, cTitle As Long
' [문자]
Private pTxtEl As Double, pTxtDim As Double, pTxtPave As Double, pTxtLine As Double
Private pFont As String, pTitleFmt As String, pDwgPrefix As String, pDwgStart As Long
Private pScaleH As String, pScaleA1 As String, pLineName As String, pTxtWF As Double
' [형상]
Private pSlope As Double
Private pBw80 As Double, pBw100 As Double, pBw160 As Double, pBwOther As Double
Private pPave02 As String, pPave04 As String
Private pHatchMin As Double, pHatchPat As String, pHatchScale As Double
Private pNegZero As Boolean, pRfilClamp As Boolean
Private pElBotOff As Long, pElTopMin As Long
Private pGrndHalf As Double, pTickHalf As Double
Private pDetailDim As Boolean, pDetailBottom As Boolean, pLeaderAll As Boolean
' [출력]
Private pWriteXdata As Boolean, pXdataApp As String
Private pShSdbH As String, pShTokjuk As String
Private pVerify As Boolean, pDbgSkipDraw As Boolean

Private pReadFail As Long

'────────────────────────── 3. CAD ──────────────────────────
Private acad As Object, doc As Object, ms As Object
Private cadAppName As String
Private gBaseX As Double, gBaseY As Double
Private gCellX As Double, gCellY As Double
Private gScale As Double

'────────────────────────── 4. 로그 ──────────────────────────
Private gErr As Collection, gWarn As Collection

'=====================================================================================
'  파라미터 기본값
'=====================================================================================
Private Sub InitParams()
    pNCol = 4: pNRow = 5
    pSheetW = 84.1: pPitchX = 18#: pPitchY = 9.5
    pOffX = 13#: pOffY = 11.5
    pPickBase = True: pBaseX = 0#: pBaseY = 0#
    pSrcGeom = "SDBH"

    pBlkTable = "관로06": pBlkBorder = "도곽": pBlkTitle = "도면번호": pBlkDot = "HDot"
    pBlkScale = 0.1: pDotScale = 0.02: pTblDX = 5#
    pBorderXref = "": pBorderW = 84.1: pBorderH = 59.4: pBorderFallback = True

    pLayCent = "CC-CENT": pLayGrnd = "CC-GRND": pLayPave = "CC-PAVE"
    pLayExcv = "CC-EXCV": pLayRfil = "CC-RFIL": pLayPipe = "CC-PIPE"
    pLayPatt = "CC-PATT": pLayLead = "CC-LEAD": pLayLeat = "CC-LEAT"
    pLayLeal = "CC-LEAL": pLayDiml = "CC-DIML": pLayText = "CC-TEXT"
    pLayTabl = "CC-TABL": pLayBorder = "CC-BORDER": pLayTitle = "CC-Title"

    cCent = 1: cGrnd = 3: cPave = 149: cExcv = 4: cRfil = 4
    cPipe = 2: cPatt = 1: cLead = 3: cLeat = 3: cLeal = 1
    cDiml = 1: cText = 3: cTabl = 3: cBorder = 7: cTitle = 7

    pTxtEl = 0.25: pTxtDim = 0.25: pTxtPave = 0.2: pTxtLine = 0.375: pTxtWF = 1#
    pFont = "굴림체": pTitleFmt = "횡 단 면 도({n})"
    pDwgPrefix = "C-": pDwgStart = 50
    pScaleH = "S = 1 : 100": pScaleA1 = "1 : 100": pLineName = "SDB"

    pSlope = 0.3
    pBw80 = 0.2: pBw100 = 0.2: pBw160 = 0#: pBwOther = 0.2
    pPave02 = "표층:0.05:CUT;기층:0.10:CUT;보조기층:0.30:SUB"
    pPave04 = "보도블럭:0.06:CUT;모래:0.04:SAND;보조기층:0.10:SUB"
    pHatchMin = 0.02: pHatchPat = "ANSI31": pHatchScale = 0.02
    pNegZero = True: pRfilClamp = True
    pElBotOff = 1: pElTopMin = 3
    pGrndHalf = 4#: pTickHalf = 0.04
    pDetailDim = True: pDetailBottom = True: pLeaderAll = False

    pWriteXdata = True: pXdataApp = "Tojuk"
    pShSdbH = "SDB(H)": pShTokjuk = "토적표"
    pVerify = True: pDbgSkipDraw = False
End Sub

'=====================================================================================
'  P 시트 생성
'=====================================================================================
Public Sub 횡단_P시트생성()
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PSHEET)
    On Error GoTo 0
    If Not ws Is Nothing Then
        If MsgBox("'" & PSHEET & "' 시트가 이미 있습니다. 기본값으로 덮어쓸까요?", _
                  vbYesNo + vbQuestion, "파라미터 시트") <> vbYes Then Exit Sub
        ws.Cells.Clear
        Application.DisplayAlerts = False
        On Error Resume Next
        ThisWorkbook.Worksheets(PSHEET).Delete
        On Error GoTo 0
        Application.DisplayAlerts = True
        Set ws = Nothing
    End If
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = PSHEET
    End If
    InitParams

    ws.Range("A1").Value = "상수관로 횡단면도 작도 파라미터  -  B열 값만 수정하세요"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 13
    ws.Range("A2:D2").Value = Array("코드", "값", "단위", "설명")
    ws.Range("A2:D2").Font.Bold = True
    ws.Range("A2:D2").Interior.Color = RGB(204, 255, 204)

    r = 3
    r = PutP(ws, r, "[배치]", "", "", "")
    r = PutP(ws, r, "N_COL", pNCol, "열", "시트당 단면 열 수 (원본 4)")
    r = PutP(ws, r, "N_ROW", pNRow, "행", "시트당 단면 행 수 (원본 5)")
    r = PutP(ws, r, "SHEET_W", pSheetW, "m", "도곽 1장 폭 = A1 x 0.1 (원본 84.1)")
    r = PutP(ws, r, "PITCH_X", pPitchX, "m", "셀 X 피치 (원본 18.0)")
    r = PutP(ws, r, "PITCH_Y", pPitchY, "m", "셀 Y 피치 (원본 9.5)")
    r = PutP(ws, r, "OFF_X", pOffX, "m", "시트 원점 -> 첫 셀 X (원본 13.0)")
    r = PutP(ws, r, "OFF_Y", pOffY, "m", "시트 원점 -> 첫 셀 Y (원본 11.5)")
    r = PutP(ws, r, "PICK_BASE", pPickBase, "", "TRUE=CAD에서 기준점 클릭 / FALSE=BASE_X,Y 사용")
    r = PutP(ws, r, "BASE_X", pBaseX, "m", "PICK_BASE=FALSE 일 때 기준 X")
    r = PutP(ws, r, "BASE_Y", pBaseY, "m", "PICK_BASE=FALSE 일 때 기준 Y")

    r = PutP(ws, r, "[입력소스]", "", "", "")
    r = PutP(ws, r, "SRC_GEOM", pSrcGeom, "", "기하 소스: SDBH(원본재현) / TOKJUK / AUTO")

    r = PutP(ws, r, "[블록]", "", "", "")
    r = PutP(ws, r, "BLK_TABLE", pBlkTable, "", "단면 표 블록명")
    r = PutP(ws, r, "BLK_BORDER", pBlkBorder, "", "도곽 블록명 (원본은 XREF)")
    r = PutP(ws, r, "BLK_TITLE", pBlkTitle, "", "도면번호(표제란) 블록명")
    r = PutP(ws, r, "BLK_DOT", pBlkDot, "", "치수 보조점 블록명")
    r = PutP(ws, r, "BLK_SCALE", pBlkScale, "", "표 블록 축척 (원본 0.1)")
    r = PutP(ws, r, "DOT_SCALE", pDotScale, "", "보조점 축척 (원본 0.02)")
    r = PutP(ws, r, "TBL_DX", pTblDX, "m", "셀 원점 -> 표 블록 X (원본 5.0)")
    r = PutP(ws, r, "BORDER_XREF", pBorderXref, "", "도곽을 XREF로 붙일 DWG 경로(비우면 블록/X선)")
    r = PutP(ws, r, "BORDER_W", pBorderW, "m", "도곽 폭 (XREF/블록 없을 때 선 작도용)")
    r = PutP(ws, r, "BORDER_H", pBorderH, "m", "도곽 높이")
    r = PutP(ws, r, "BORDER_FALLBACK", pBorderFallback, "", "도곽블록이 없으면 외곽선을 직접 작도")

    r = PutP(ws, r, "[레이어]", "", "", "")
    r = PutP(ws, r, "LAY_CENT", pLayCent, "", "중심틱")
    r = PutP(ws, r, "LAY_GRND", pLayGrnd, "", "지반선")
    r = PutP(ws, r, "LAY_PAVE", pLayPave, "", "포장층 경계")
    r = PutP(ws, r, "LAY_EXCV", pLayExcv, "", "굴착선")
    r = PutP(ws, r, "LAY_RFIL", pLayRfil, "", "관저 기준선")
    r = PutP(ws, r, "LAY_PIPE", pLayPipe, "", "관원")
    r = PutP(ws, r, "LAY_PATT", pLayPatt, "", "해치")
    r = PutP(ws, r, "LAY_LEAD", pLayLead, "", "상세치수 보조선")
    r = PutP(ws, r, "LAY_LEAT", pLayLeat, "", "상세치수 문자")
    r = PutP(ws, r, "LAY_LEAL", pLayLeal, "", "포장 리더")
    r = PutP(ws, r, "LAY_DIML", pLayDiml, "", "고도축")
    r = PutP(ws, r, "LAY_TEXT", pLayText, "", "고도/라인명 문자")
    r = PutP(ws, r, "LAY_TABL", pLayTabl, "", "표 블록")
    r = PutP(ws, r, "LAY_BORDER", pLayBorder, "", "도곽")
    r = PutP(ws, r, "LAY_TITLE", pLayTitle, "", "표제란")

    r = PutP(ws, r, "[색상 ACI]", "", "", "")
    r = PutP(ws, r, "COL_CENT", cCent, "", "1=빨강 2=노랑 3=초록 4=청록 7=흰색")
    r = PutP(ws, r, "COL_GRND", cGrnd, "", "")
    r = PutP(ws, r, "COL_PAVE", cPave, "", "149")
    r = PutP(ws, r, "COL_EXCV", cExcv, "", "")
    r = PutP(ws, r, "COL_RFIL", cRfil, "", "")
    r = PutP(ws, r, "COL_PIPE", cPipe, "", "")
    r = PutP(ws, r, "COL_PATT", cPatt, "", "")
    r = PutP(ws, r, "COL_LEAD", cLead, "", "")
    r = PutP(ws, r, "COL_LEAT", cLeat, "", "")
    r = PutP(ws, r, "COL_LEAL", cLeal, "", "")
    r = PutP(ws, r, "COL_DIML", cDiml, "", "")
    r = PutP(ws, r, "COL_TEXT", cText, "", "")
    r = PutP(ws, r, "COL_TABL", cTabl, "", "")
    r = PutP(ws, r, "COL_BORDER", cBorder, "", "")
    r = PutP(ws, r, "COL_TITLE", cTitle, "", "")

    r = PutP(ws, r, "[문자]", "", "", "")
    r = PutP(ws, r, "TXT_EL", pTxtEl, "m", "고도 문자 높이 (원본 0.25)")
    r = PutP(ws, r, "TXT_DIM", pTxtDim, "m", "H=/B= 문자 높이 (원본 0.25)")
    r = PutP(ws, r, "TXT_PAVE", pTxtPave, "m", "포장 리더 문자 (원본 0.20)")
    r = PutP(ws, r, "TXT_LINE", pTxtLine, "m", "라인명 문자 (원본 0.375)")
    r = PutP(ws, r, "TXT_WF", pTxtWF, "", "문자 폭계수")
    r = PutP(ws, r, "FONT", pFont, "", "문자 스타일 (없으면 Standard)")
    r = PutP(ws, r, "TITLE_FMT", pTitleFmt, "", "도면명 형식 : {n} 이 번호로 치환")
    r = PutP(ws, r, "DWG_PREFIX", pDwgPrefix, "", "도면번호 접두 (원본 C-)")
    r = PutP(ws, r, "DWG_START", pDwgStart, "", "도면번호 시작값 (원본 50)")
    r = PutP(ws, r, "SCALE_H", pScaleH, "", "축척H 문자")
    r = PutP(ws, r, "SCALE_A1", pScaleA1, "", "축척A1 문자")
    r = PutP(ws, r, "LINE_NAME", pLineName, "", "라인명 (원본 SDB)")

    r = PutP(ws, r, "[형상]", "", "", "")
    r = PutP(ws, r, "SLOPE", pSlope, "", "굴착 측면 경사 (원본 0.3)")
    r = PutP(ws, r, "BW_80", pBw80, "m", "관기초 편측폭 D80 (원본 0.20)")
    r = PutP(ws, r, "BW_100", pBw100, "m", "관기초 편측폭 D100 (원본 0.20)")
    r = PutP(ws, r, "BW_160", pBw160, "m", "관기초 편측폭 D160 (실데이터 0 / 절차서 0.20)")
    r = PutP(ws, r, "BW_OTHER", pBwOther, "m", "그 외 관경 편측폭")
    r = PutP(ws, r, "PAVE_02", pPave02, "", "OP02 구성  이름:두께:구분  구분=CUT/SUB/SAND")
    r = PutP(ws, r, "PAVE_04", pPave04, "", "OP04 구성")
    r = PutP(ws, r, "HATCH_MIN", pHatchMin, "sqm", "해치 작도 최소 되메우기(관주위)")
    r = PutP(ws, r, "HATCH_PAT", pHatchPat, "", "해치 패턴 (원본 ANSI31)")
    r = PutP(ws, r, "HATCH_SCALE", pHatchScale, "", "해치 축척 (원본 0.02)")
    r = PutP(ws, r, "NEG_TO_ZERO", pNegZero, "", "되메우기 음수를 0으로")
    r = PutP(ws, r, "RFIL_CLAMP", pRfilClamp, "", "관저선이 굴착바닥을 넘지 않게 클램프")
    r = PutP(ws, r, "EL_BOT_OFF", pElBotOff, "m", "고도축 하단 = floor(INV) - 값 (원본 1)")
    r = PutP(ws, r, "EL_TOP_MIN", pElTopMin, "m", "고도축 상단 = 하단 + 값 + floor(H) (원본 3)")
    r = PutP(ws, r, "GRND_HALF", pGrndHalf, "m", "지반선 반폭 (원본 4.0)")
    r = PutP(ws, r, "TICK_HALF", pTickHalf, "m", "중심틱 반높이 (원본 0.04)")
    r = PutP(ws, r, "DETAIL_DIM", pDetailDim, "", "상세치수(H/B) 작도")
    r = PutP(ws, r, "DETAIL_BOTTOM", pDetailBottom, "", "저부 관경/기초폭 치수 작도")
    r = PutP(ws, r, "LEADER_ALL", pLeaderAll, "", "TRUE=모든 단면에 포장리더 / FALSE=첫단면+포장변경")

    r = PutP(ws, r, "[출력]", "", "", "")
    r = PutP(ws, r, "WRITE_XDATA", pWriteXdata, "", "관로06 에 XDATA 기록")
    r = PutP(ws, r, "XDATA_APP", pXdataApp, "", "XDATA AppID (원본 Tojuk)")
    r = PutP(ws, r, "SH_SDBH", pShSdbH, "", "횡단 입력 시트명")
    r = PutP(ws, r, "SH_TOKJUK", pShTokjuk, "", "토적표 시트명")
    r = PutP(ws, r, "VERIFY", pVerify, "", "작도 후 검증 시트 갱신")
    r = PutP(ws, r, "DBG_SKIP_DRAW", pDbgSkipDraw, "", "TRUE=작도 없이 검증만 (테스트)")

    With ws
        .Columns("A").ColumnWidth = 18
        .Columns("B").ColumnWidth = 40
        .Columns("C").ColumnWidth = 8
        .Columns("D").ColumnWidth = 66
        .Range(.Cells(3, 2), .Cells(r, 2)).HorizontalAlignment = xlLeft
        .Range(.Cells(1, 1), .Cells(r, 4)).VerticalAlignment = xlTop
    End With
    MsgBox "'" & PSHEET & "' 시트를 만들었습니다." & vbLf & vbLf & _
           "1) B열 값을 확인/수정" & vbLf & _
           "2) [횡단_1단계_검증] 실행" & vbLf & _
           "3) [횡단_작도] 실행", vbInformation, "완료"
End Sub

Private Function PutP(ws As Worksheet, r As Long, k As String, v As Variant, _
                       u As String, desc As String) As Long
    ws.Cells(r, 1).Value = k
    If Left$(k, 1) = "[" Then
        With ws.Range(ws.Cells(r, 1), ws.Cells(r, 4))
            .Interior.Color = RGB(226, 234, 246)
        End With
        ws.Cells(r, 1).Font.Bold = True
    Else
        ws.Cells(r, 2).Value = v
        ws.Cells(r, 3).Value = u
        ws.Cells(r, 4).Value = desc
    End If
    PutP = r + 1
End Function
'=====================================================================================
'  P 시트 -> 변수 읽기
'=====================================================================================
Public Sub 횡단_파라미터읽기()
    ReadParams
    MsgBox "'" & PSHEET & "' 시트 값을 읽었습니다. (읽기실패 " & pReadFail & "건)", _
           vbInformation, "파라미터"
End Sub

Private Sub ReadParams()
    Dim ws As Worksheet, r As Long, lastR As Long, k As String, v As Variant
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PSHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        InitParams
        Exit Sub
    End If
    pReadFail = 0
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 3 To lastR
        k = Trim$(CStr(ws.Cells(r, 1).Value & ""))
        If Len(k) > 0 And Left$(k, 1) <> "[" Then
            v = ws.Cells(r, 2).Value
            ApplyOne k, v
        End If
    Next r
    ' 값이 비어 있으면 적용되지 않으므로, 미적용 항목은 InitParams 기본값이 남는다.
    If pNCol < 1 Then pNCol = 1
    If pNRow < 1 Then pNRow = 1
    If pBlkScale <= 0 Then pBlkScale = 0.1
    If pDotScale <= 0 Then pDotScale = 0.02
    If pSheetW <= 0 Then pSheetW = 84.1
    If pPitchX <= 0 Then pPitchX = 18#
    If pPitchY <= 0 Then pPitchY = 9.5
End Sub

' v 가 Empty/Nothing 이면 기본값 유지(건너뜀)
Private Sub ApplyOne(ByVal k As String, ByVal v As Variant)
    Select Case UCase$(k)
    ' ---- 배치 ----
    Case "N_COL":         pNCol = CLng(Num(v, pNCol))
    Case "N_ROW":         pNRow = CLng(Num(v, pNRow))
    Case "SHEET_W":       pSheetW = Num(v, pSheetW)
    Case "PITCH_X":       pPitchX = Num(v, pPitchX)
    Case "PITCH_Y":       pPitchY = Num(v, pPitchY)
    Case "OFF_X":         pOffX = Num(v, pOffX)
    Case "OFF_Y":         pOffY = Num(v, pOffY)
    Case "PICK_BASE":     pPickBase = Boo(v, pPickBase)
    Case "BASE_X":        pBaseX = Num(v, pBaseX)
    Case "BASE_Y":        pBaseY = Num(v, pBaseY)
    ' ---- 소스 ----
    Case "SRC_GEOM":      pSrcGeom = UCase$(Trim$(CStr(v & "")))
    ' ---- 블록 ----
    Case "BLK_TABLE":     pBlkTable = Trim$(CStr(v & ""))
    Case "BLK_BORDER":    pBlkBorder = Trim$(CStr(v & ""))
    Case "BLK_TITLE":     pBlkTitle = Trim$(CStr(v & ""))
    Case "BLK_DOT":       pBlkDot = Trim$(CStr(v & ""))
    Case "BLK_SCALE":     pBlkScale = Num(v, pBlkScale)
    Case "DOT_SCALE":     pDotScale = Num(v, pDotScale)
    Case "TBL_DX":        pTblDX = Num(v, pTblDX)
    Case "BORDER_XREF":   pBorderXref = Trim$(CStr(v & ""))
    Case "BORDER_W":      pBorderW = Num(v, pBorderW)
    Case "BORDER_H":      pBorderH = Num(v, pBorderH)
    Case "BORDER_FALLBACK": pBorderFallback = Boo(v, pBorderFallback)
    ' ---- 레이어 ----
    Case "LAY_CENT":      pLayCent = Trim$(CStr(v & ""))
    Case "LAY_GRND":      pLayGrnd = Trim$(CStr(v & ""))
    Case "LAY_PAVE":      pLayPave = Trim$(CStr(v & ""))
    Case "LAY_EXCV":      pLayExcv = Trim$(CStr(v & ""))
    Case "LAY_RFIL":      pLayRfil = Trim$(CStr(v & ""))
    Case "LAY_PIPE":      pLayPipe = Trim$(CStr(v & ""))
    Case "LAY_PATT":      pLayPatt = Trim$(CStr(v & ""))
    Case "LAY_LEAD":      pLayLead = Trim$(CStr(v & ""))
    Case "LAY_LEAT":      pLayLeat = Trim$(CStr(v & ""))
    Case "LAY_LEAL":      pLayLeal = Trim$(CStr(v & ""))
    Case "LAY_DIML":      pLayDiml = Trim$(CStr(v & ""))
    Case "LAY_TEXT":      pLayText = Trim$(CStr(v & ""))
    Case "LAY_TABL":      pLayTabl = Trim$(CStr(v & ""))
    Case "LAY_BORDER":    pLayBorder = Trim$(CStr(v & ""))
    Case "LAY_TITLE":     pLayTitle = Trim$(CStr(v & ""))
    ' ---- 색상 ----
    Case "COL_CENT":      cCent = CLng(Num(v, cCent))
    Case "COL_GRND":      cGrnd = CLng(Num(v, cGrnd))
    Case "COL_PAVE":      cPave = CLng(Num(v, cPave))
    Case "COL_EXCV":      cExcv = CLng(Num(v, cExcv))
    Case "COL_RFIL":      cRfil = CLng(Num(v, cRfil))
    Case "COL_PIPE":      cPipe = CLng(Num(v, cPipe))
    Case "COL_PATT":      cPatt = CLng(Num(v, cPatt))
    Case "COL_LEAD":      cLead = CLng(Num(v, cLead))
    Case "COL_LEAT":      cLeat = CLng(Num(v, cLeat))
    Case "COL_LEAL":      cLeal = CLng(Num(v, cLeal))
    Case "COL_DIML":      cDiml = CLng(Num(v, cDiml))
    Case "COL_TEXT":      cText = CLng(Num(v, cText))
    Case "COL_TABL":      cTabl = CLng(Num(v, cTabl))
    Case "COL_BORDER":    cBorder = CLng(Num(v, cBorder))
    Case "COL_TITLE":     cTitle = CLng(Num(v, cTitle))
    ' ---- 문자 ----
    Case "TXT_EL":        pTxtEl = Num(v, pTxtEl)
    Case "TXT_DIM":       pTxtDim = Num(v, pTxtDim)
    Case "TXT_PAVE":      pTxtPave = Num(v, pTxtPave)
    Case "TXT_LINE":      pTxtLine = Num(v, pTxtLine)
    Case "TXT_WF":        pTxtWF = Num(v, pTxtWF)
    Case "FONT":          pFont = Trim$(CStr(v & ""))
    Case "TITLE_FMT":     pTitleFmt = CStr(v & "")
    Case "DWG_PREFIX":    pDwgPrefix = CStr(v & "")
    Case "DWG_START":     pDwgStart = CLng(Num(v, pDwgStart))
    Case "SCALE_H":       pScaleH = CStr(v & "")
    Case "SCALE_A1":      pScaleA1 = CStr(v & "")
    Case "LINE_NAME":     pLineName = CStr(v & "")
    ' ---- 형상 ----
    Case "SLOPE":         pSlope = Num(v, pSlope)
    Case "BW_80":         pBw80 = Num(v, pBw80)
    Case "BW_100":        pBw100 = Num(v, pBw100)
    Case "BW_160":        pBw160 = Num(v, pBw160)
    Case "BW_OTHER":      pBwOther = Num(v, pBwOther)
    Case "PAVE_02":       pPave02 = CStr(v & "")
    Case "PAVE_04":       pPave04 = CStr(v & "")
    Case "HATCH_MIN":     pHatchMin = Num(v, pHatchMin)
    Case "HATCH_PAT":     pHatchPat = Trim$(CStr(v & ""))
    Case "HATCH_SCALE":   pHatchScale = Num(v, pHatchScale)
    Case "NEG_TO_ZERO":   pNegZero = Boo(v, pNegZero)
    Case "RFIL_CLAMP":    pRfilClamp = Boo(v, pRfilClamp)
    Case "EL_BOT_OFF":    pElBotOff = CLng(Num(v, pElBotOff))
    Case "EL_TOP_MIN":    pElTopMin = CLng(Num(v, pElTopMin))
    Case "GRND_HALF":     pGrndHalf = Num(v, pGrndHalf)
    Case "TICK_HALF":     pTickHalf = Num(v, pTickHalf)
    Case "DETAIL_DIM":    pDetailDim = Boo(v, pDetailDim)
    Case "DETAIL_BOTTOM": pDetailBottom = Boo(v, pDetailBottom)
    Case "LEADER_ALL":    pLeaderAll = Boo(v, pLeaderAll)
    ' ---- 출력 ----
    Case "WRITE_XDATA":   pWriteXdata = Boo(v, pWriteXdata)
    Case "XDATA_APP":     pXdataApp = Trim$(CStr(v & ""))
    Case "SH_SDBH":       pShSdbH = Trim$(CStr(v & ""))
    Case "SH_TOKJUK":     pShTokjuk = Trim$(CStr(v & ""))
    Case "VERIFY":        pVerify = Boo(v, pVerify)
    Case "DBG_SKIP_DRAW": pDbgSkipDraw = Boo(v, pDbgSkipDraw)
    Case Else:            pReadFail = pReadFail + 1
    End Select
End Sub

Private Function Num(ByVal v As Variant, ByVal dflt As Double) As Double
    If IsEmpty(v) Then Num = dflt: Exit Function
    If VarType(v) = vbNull Then Num = dflt: Exit Function
    If IsNumeric(v) Then Num = CDbl(v) Else Num = dflt
End Function

Private Function Boo(ByVal v As Variant, ByVal dflt As Boolean) As Boolean
    Dim s As String
    If IsEmpty(v) Then Boo = dflt: Exit Function
    If VarType(v) = vbBoolean Then Boo = CBool(v): Exit Function
    s = UCase$(Trim$(CStr(v & "")))
    Select Case s
    Case "TRUE", "1", "Y", "YES", "O", "예": Boo = True
    Case "FALSE", "0", "N", "NO", "X", "아니오": Boo = False
    Case Else: Boo = dflt
    End Select
End Function

'=====================================================================================
'  로그
'=====================================================================================
Private Sub LogInit()
    Set gErr = New Collection
    Set gWarn = New Collection
End Sub

Private Sub AddErr(ByVal s As String)
    If gErr Is Nothing Then Exit Sub
    On Error Resume Next
    gErr.Add s
End Sub

Private Sub AddWarn(ByVal s As String)
    If gWarn Is Nothing Then Exit Sub
    On Error Resume Next
    gWarn.Add s
End Sub

'=====================================================================================
'  CAD 접속  ★ AUTOCAD / ZWCAD / GstarCAD / BricsCAD 를 모두 시도
'  ------------------------------------------------------------------------------
'  1) 이미 실행 중인 캐드에 붙는다(GetObject)  -> 사용자가 열어둔 도면에 작도
'  2) 없으면 새로 실행한다(CreateObject)
'  3) ZWCAD 를 AutoCAD 보다 먼저 시도한다(국내 관급설계 배포본 우선)
'=====================================================================================
Public Function xcConnectCAD() As Boolean
    Dim cand As Variant, i As Long, pid As String
    Dim found As Object
    xcConnectCAD = False
    cadAppName = ""

    ' ZWCAD 계열 -> AutoCAD 계열 -> 기타 순
    cand = Array( _
        "ZWCAD.Application", "ZWCAD.Application.2026", "ZWCAD.Application.2025", _
        "ZWCAD.Application.2024", "ZWCAD.Application.2023", "ZWCAD.Application.2022", _
        "ZWCAD.Application.2021", "ZWCAD.Application.2020", "ZWCAD.Application.2019", _
        "ZWCAD.Application.2018", "ZWCAD.Application.2017", _
        "AutoCAD.Application", _
        "AutoCAD.Application.26.0", "AutoCAD.Application.25.1", "AutoCAD.Application.25.0", _
        "AutoCAD.Application.24.3", "AutoCAD.Application.24.2", "AutoCAD.Application.24.1", _
        "AutoCAD.Application.24.0", "AutoCAD.Application.23.1", "AutoCAD.Application.23.0", _
        "AutoCAD.Application.22.0", "AutoCAD.Application.21.0", "AutoCAD.Application.20.1", _
        "AutoCAD.Application.20.0", "AutoCAD.Application.19.1", "AutoCAD.Application.19.0", _
        "AutoCAD.Application.18.2", "AutoCAD.Application.18.1", "AutoCAD.Application.18.0", _
        "AutoCAD.Application.17.2", "AutoCAD.Application.17.1", "AutoCAD.Application.17.0", _
        "AutoCAD.Application.16.1", "AutoCAD.Application.16.0", _
        "GstarCAD.Application", "GstarCAD.Application.2025", "GstarCAD.Application.2024", _
        "BricscadApp.AcadApplication")

    For i = LBound(cand) To UBound(cand)
        pid = CStr(cand(i))
        Set found = Nothing
        On Error Resume Next
        Set found = GetObject(, pid)                 ' 실행 중인 캐드
        If found Is Nothing Then
            Err.Clear
            Set found = CreateObject(pid)            ' 새로 실행
        End If
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        If Not found Is Nothing Then
            Set acad = found
            Exit For
        End If
    Next i

    If acad Is Nothing Then
        MsgBox "캐드(AutoCAD / ZWCAD / GstarCAD / BricsCAD)를 찾지 못했습니다." & vbLf & vbLf & _
               "· 캐드를 먼저 실행한 뒤 다시 시도하세요." & vbLf & _
               "· 32/64bit 또는 권한 문제라면 캐드를 '관리자 권한'으로 실행해 보세요.", _
               vbExclamation, "CAD 접속 실패"
        Exit Function
    End If

    On Error Resume Next
    cadAppName = acad.Caption
    If Err.Number <> 0 Then cadAppName = pid: Err.Clear
    On Error GoTo 0

    ' 도면이 하나도 없으면 새로 만든다
    On Error Resume Next
    If acad.Documents.Count = 0 Then acad.Documents.Add
    Set doc = acad.ActiveDocument
    If Err.Number <> 0 Then
        Err.Clear
        Set doc = acad.Documents.Add
    End If
    On Error GoTo 0

    If doc Is Nothing Then
        MsgBox "캐드 도면(Document)을 열지 못했습니다.", vbExclamation, "CAD 접속 실패"
        Exit Function
    End If

    Set ms = doc.ModelSpace

    ' 단위/표시 설정 (실패해도 계속)
    On Error Resume Next
    doc.SetVariable "INSUNITS", 6      ' meter
    doc.SetVariable "LUNITS", 2
    doc.SetVariable "LUPREC", 4
    If pXdataApp <> "" Then
        Dim ra As Object
        Set ra = doc.RegAppTable.Item(pXdataApp)
        If ra Is Nothing Then doc.RegAppTable.Add pXdataApp
    End If
    On Error GoTo 0

    xcConnectCAD = True
End Function

' 문자 스타일 확보 (없으면 만든다)
Private Function EnsureTextStyle(ByVal nm As String) As String
    Dim st As Object
    EnsureTextStyle = "Standard"
    If Len(nm) = 0 Then Exit Function
    On Error Resume Next
    Set st = doc.TextStyles.Item(nm)
    If st Is Nothing Then
        Set st = doc.TextStyles.Add(nm)
        st.fontFile = "simplex.shx"
        st.BigFontFile = "hfont.shx"
    End If
    If Err.Number <> 0 Then Err.Clear: Exit Function
    On Error GoTo 0
    EnsureTextStyle = nm
End Function

'=====================================================================================
'  레이어 생성
'=====================================================================================
Private Function EnsureLayer(ByVal nm As String, ByVal col As Long) As String
    Dim la As Object
    EnsureLayer = "0"
    If Len(nm) = 0 Then Exit Function
    On Error Resume Next
    Set la = doc.Layers.Item(nm)
    If la Is Nothing Then
        Set la = doc.Layers.Add(nm)
        la.color = col
    End If
    If Err.Number <> 0 Then Err.Clear: Exit Function
    On Error GoTo 0
    EnsureLayer = nm
End Function

Private Sub SetupLayers()
    EnsureLayer pLayCent, cCent
    EnsureLayer pLayGrnd, cGrnd
    EnsureLayer pLayPave, cPave
    EnsureLayer pLayExcv, cExcv
    EnsureLayer pLayRfil, cRfil
    EnsureLayer pLayPipe, cPipe
    EnsureLayer pLayPatt, cPatt
    EnsureLayer pLayLead, cLead
    EnsureLayer pLayLeat, cLeat
    EnsureLayer pLayLeal, cLeal
    EnsureLayer pLayDiml, cDiml
    EnsureLayer pLayText, cText
    EnsureLayer pLayTabl, cTabl
    EnsureLayer pLayBorder, cBorder
    EnsureLayer pLayTitle, cTitle
    EnsureTextStyle pFont
    EnsureTextStyle "TITLE"
End Sub

'=====================================================================================
'  원시 작도 함수
'=====================================================================================
Private Function xcP3(ByVal x As Double, ByVal y As Double) As Variant
    xcP3 = Array(x, y, 0#)
End Function

Private Function xcLine(ByVal x1 As Double, ByVal y1 As Double, _
                         ByVal x2 As Double, ByVal y2 As Double, _
                         ByVal lay As String) As Object
    Dim e As Object
    Set xcLine = Nothing
    On Error Resume Next
    Set e = ms.AddLine(xcP3(x1, y1), xcP3(x2, y2))
    If Err.Number <> 0 Then Err.Clear: Exit Function
    On Error GoTo 0
    If Len(lay) > 0 Then e.Layer = lay
    Set xcLine = e
End Function

' pts = "x1,y1;x2,y2;..."  (2D)
Private Function xcPoly(ByVal pts As String, ByVal lay As String, _
                         Optional ByVal bClosed As Boolean = False) As Object
    Dim a() As Double, f As Variant, i As Long, m As Long
    Dim e As Object
    Set xcPoly = Nothing
    f = Split(pts, ";")
    m = -1
    ReDim a(0 To (UBound(f) + 1) * 2 - 1)
    For i = LBound(f) To UBound(f)
        Dim xy As Variant
        xy = Split(f(i), ",")
        If UBound(xy) >= 1 Then
            m = m + 1
            a(m * 2) = CDbl(xy(0))
            a(m * 2 + 1) = CDbl(xy(1))
        End If
    Next i
    If m < 1 Then Exit Function
    On Error Resume Next
    Set e = ms.AddLightWeightPolyline(a)
    If Err.Number <> 0 Then Err.Clear: Exit Function
    On Error GoTo 0
    If bClosed Then e.Closed = True
    If Len(lay) > 0 Then e.Layer = lay
    Set xcPoly = e
End Function

' 원시 좌표 배열로 폴리라인
Private Function xcPolyA(ByRef a() As Double, ByVal n As Long, ByVal lay As String, _
                          Optional ByVal bClosed As Boolean = False) As Object
    Dim e As Object, a2() As Double, i As Long
    Set xcPolyA = Nothing
    If n < 2 Then Exit Function
    ReDim a2(0 To n * 2 - 1)
    For i = 0 To n - 1
        a2(i * 2) = a(i * 2)
        a2(i * 2 + 1) = a(i * 2 + 1)
    Next i
    On Error Resume Next
    Set e = ms.AddLightWeightPolyline(a2)
    If Err.Number <> 0 Then Err.Clear: Exit Function
    On Error GoTo 0
    If bClosed Then e.Closed = True
    If Len(lay) > 0 Then e.Layer = lay
    Set xcPolyA = e
End Function

Private Function xcText(ByVal s As String, ByVal x As Double, ByVal y As Double, _
                         ByVal h As Double, ByVal lay As String, _
                         Optional ByVal rot As Double = 0#, _
                         Optional ByVal al As Long = 0, _
                         Optional ByVal rotGiven As Boolean = True) As Object
    Dim e As Object
    Set xcText = Nothing
    On Error Resume Next
    Set e = ms.AddText(s, xcP3(x, y), h)
    If Err.Number <> 0 Then Err.Clear: Exit Function
    On Error GoTo 0
    If Len(lay) > 0 Then e.Layer = lay
    If pTxtWF <> 1 Then
        On Error Resume Next
        e.ScaleFactor = pTxtWF
        On Error GoTo 0
    End If
    ' 회전각(도) : DXF 는 라디안이 아니라 도 단위
    If rotGiven And Abs(rot) > 1E-12 Then
        On Error Resume Next
        e.Rotation = rot * PI / 180#
        On Error GoTo 0
    End If
    If al <> AL_LEFT Then
        On Error Resume Next
        e.Alignment = al
        e.TextAlignmentPoint = xcP3(x, y)
        On Error GoTo 0
    End If
    Set xcText = e
End Function

Private Function xcCircle(ByVal cx As Double, ByVal cy As Double, ByVal r As Double, _
                           ByVal lay As String) As Object
    Dim e As Object
    Set xcCircle = Nothing
    If r <= 0 Then Exit Function
    On Error Resume Next
    Set e = ms.AddCircle(xcP3(cx, cy), r)
    If Err.Number <> 0 Then Err.Clear: Exit Function
    On Error GoTo 0
    If Len(lay) > 0 Then e.Layer = lay
    Set xcCircle = e
End Function

'=====================================================================================
'  10. 배치 (셀 원점 · 시트 원점)
'=====================================================================================
Private Sub SheetOrigin(ByVal k As Long, ByRef ox As Double, ByRef oy As Double)
    Dim sh As Long
    sh = k \ (pNCol * pNRow)
    ox = gBaseX + sh * pSheetW
    oy = gBaseY
End Sub

Private Sub CellXY(ByVal k As Long, ByRef cx As Double, ByRef cy As Double)
    Dim sh As Long, rem20 As Long, col As Long, row As Long
    sh = k \ (pNCol * pNRow)
    rem20 = k - sh * (pNCol * pNRow)
    col = rem20 \ pNRow
    row = rem20 - col * pNRow
    cx = gBaseX + sh * pSheetW + pOffX + col * pPitchX
    cy = gBaseY + pOffY + row * pPitchY
End Sub

Private Function IsFirstOfSheet(ByVal k As Long) As Boolean
    IsFirstOfSheet = ((k Mod (pNCol * pNRow)) = 0)
End Function

Private Function SheetNo(ByVal k As Long) As Long
    SheetNo = k \ (pNCol * pNRow) + 1
End Function

'=====================================================================================
'  11. 단면 1개 작도
'     좌표계 : 셀 원점 = (0,0), 지반면 = y 0, 아래가 음수
'=====================================================================================
Private Sub DrawSection(ByVal k As Long)
    Dim cx As Double, cy As Double
    Dim i As Long, z As Double
    Dim e As Object, circ As Object, poly As Object
    Dim pts As String

    Call CellXY(k, cx, cy)
    SecCalc k

    ' ---- 1) 중심틱 ----
    xcLine cx, cy + pTickHalf, cx, cy - pTickHalf, pLayCent

    ' ---- 2) 지반선 (좌/우 2개) ----
    xcPoly cx & "," & cy & ";" & (cx - pGrndHalf) & "," & cy, pLayGrnd
    xcPoly cx & "," & cy & ";" & (cx + pGrndHalf) & "," & cy, pLayGrnd

    ' ---- 3) 포장층 경계 ----
    z = 0#
    For i = 0 To gNLay - 1
        z = z + gLayTh(i)
        xcPoly cx & "," & (cy - z) & ";" & (cx - pGrndHalf) & "," & (cy - z), pLayPave
        xcPoly cx & "," & (cy - z) & ";" & (cx + pGrndHalf) & "," & (cy - z), pLayPave
    Next i

    ' ---- 4) 굴착 폴리곤 ----
    '  원본과 동일하게 closed 플래그를 쓰지 않고 시작점=끝점을 중복시켜 닫는다.
    If gHasExc Then
        pts = Pt(cx + 0#, cy - gH)
        pts = pts & ";" & Pt(cx - ghb, cy - gH)
        pts = pts & ";" & Pt(cx - ghb, cy - gH)
        pts = pts & ";" & Pt(cx - ghb, cy - gH)
        pts = pts & ";" & Pt(cx - ght, cy - gtv)
        pts = pts & ";" & Pt(cx - ght, cy)
        pts = pts & ";" & Pt(cx + 0#, cy)
        pts = pts & ";" & Pt(cx + 0#, cy)
        pts = pts & ";" & Pt(cx + ght, cy)
        pts = pts & ";" & Pt(cx + ght, cy - gtv)
        pts = pts & ";" & Pt(cx + ghb, cy - gH)
        pts = pts & ";" & Pt(cx + ghb, cy - gH)
        pts = pts & ";" & Pt(cx + ghb, cy - gH)
        pts = pts & ";" & Pt(cx + 0#, cy - gH)
        Set poly = xcPoly(pts, pLayExcv, False)
    End If

    ' ---- 5) 관원 ----
    Set circ = xcCircle(cx, cy - (gH - gD / 2#), gD / 2#, pLayPipe)

    ' ---- 6) 관저 기준선 (RFIL) ----
    If gHasExc And gzRfil > 0# Then
        xcLine cx - gRfilHalf, cy - gzRfil, cx + gRfilHalf, cy - gzRfil, pLayRfil
    End If

    ' ---- 7) 해치 (관주위 되메우기 영역) ----
    If gHasExc And gAr >= pHatchMin And gzRfil > TINY Then
        HatchBackfill cx, cy, circ
    End If

    ' ---- 8) 표 블록 ----
    DrawTableBlock cx, cy, k

    ' ---- 9) 고도축 ----
    DrawElAxis cx, cy, k

    ' ---- 10) 상세치수 / 포장리더 ----
    If pDetailDim And IsFirstOfSheet(k) And gHasExc Then DrawDetailDim cx, cy
    If NeedPaveLeader(k) Then DrawPaveLeader cx, cy
End Sub

Private Function Pt(ByVal x As Double, ByVal y As Double) As String
    Pt = Trim$(Str$(x)) & "," & Trim$(Str$(y))
End Function

' 포장리더를 그릴 단면인가 : (첫단면 or 포장변경) and 포장층 있음
Private Function NeedPaveLeader(ByVal k As Long) As Boolean
    Dim first As Boolean, chg As Boolean
    If gNLay = 0 Then NeedPaveLeader = False: Exit Function
    first = IsFirstOfSheet(k)
    chg = False
    If pLeaderAll Then
        chg = True
    ElseIf k > 0 Then
        chg = (PaveCode(sCode(k)) <> PaveCode(sCode(k - 1)))
    End If
    NeedPaveLeader = (first Or chg)
End Function

'=====================================================================================
'  5. 수치 유틸
'=====================================================================================
' 반올림 : 0.5 는 0 에서 먼 쪽으로 (엑셀 ROUND 와 동일)
Private Function Rnd2(ByVal x As Double) As Double
    Dim s As Double, a As Double
    If x = 0# Then Rnd2 = 0#: Exit Function
    s = 1#: If x < 0# Then s = -1#
    a = Abs(x)
    Rnd2 = s * Int(a * 100# + 0.5 + 0.0000001) / 100#
End Function

Private Function Dv(ByVal v As Variant) As Double
    Dv = 0#
    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    Dv = CDbl(v)
End Function

' 표 블록 속성값 : 0(또는 미입력)은 빈칸, 그 외 소수 2자리
Private Function F2(ByVal v As Variant) As String
    Dim d As Double
    F2 = ""
    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    d = CDbl(v)
    If Abs(d) < 0.005 Then Exit Function
    F2 = Format$(d, "0.00")
End Function

' XDATA 값 : 정수는 정수, 실수는 소수 2자리, 미입력은 빈칸
Private Function FI(ByVal v As Variant) As String
    Dim d As Double
    FI = ""
    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If Not IsNumeric(v) Then FI = CStr(v): Exit Function
    d = CDbl(v)
    If Abs(d - Int(d)) < 0.000001 Then
        FI = Format$(d, "0")
    Else
        FI = Format$(d, "0.00")
    End If
End Function

'=====================================================================================
'  6. 시트 찾기 + 입력 읽기
'=====================================================================================
Private Function SheetByName(ByVal nm As String) As Worksheet
    Set SheetByName = Nothing
    On Error Resume Next
    Set SheetByName = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
End Function

Private Function FindTokjuk() As Worksheet
    Dim ws As Worksheet, c As Long, t As String
    Set FindTokjuk = Nothing
    Set ws = SheetByName(pShTokjuk)
    If Not ws Is Nothing Then Set FindTokjuk = ws: Exit Function
    For c = 1 To ThisWorkbook.Worksheets.Count
        t = ThisWorkbook.Worksheets(c).Name
        If InStr(t, "토적") > 0 Then
            Set FindTokjuk = ThisWorkbook.Worksheets(c)
            pShTokjuk = t
            Exit Function
        End If
    Next c
End Function

Private Function FindSdbH() As Worksheet
    Dim ws As Worksheet, c As Long, t As String
    Set FindSdbH = Nothing
    Set ws = SheetByName(pShSdbH)
    If Not ws Is Nothing Then Set FindSdbH = ws: Exit Function
    For c = 1 To ThisWorkbook.Worksheets.Count
        t = ThisWorkbook.Worksheets(c).Name
        If InStr(UCase$(t), "SDB(H)") > 0 Then
            Set FindSdbH = ThisWorkbook.Worksheets(c)
            pShSdbH = t
            Exit Function
        End If
    Next c
End Function

' 토적표 : 6행부터 측점(B열)이 있는 행까지 = 단면 1개
Private Function ReadTokjuk() As Long
    Dim ws As Worksheet, v As Variant, i As Long, r As Long, n As Long
    ReDim sSta(0 To 0): ReDim sGH(0 To 0): ReDim sInv(0 To 0)
    ReDim sDia(0 To 0): ReDim sCode(0 To 0): ReDim sSlope(0 To 0)
    ReDim sExc(0 To 0): ReDim sWet(0 To 0): ReDim sAr(0 To 0): ReDim sAb(0 To 0)
    ReDim sAsp(0 To 0): ReDim sCon(0 To 0): ReDim sAsc(0 To 0): ReDim sBodo(0 To 0)
    ReDim sSand(0 To 0): ReDim sSub(0 To 0): ReDim sFrost(0 To 0)
    ReDim sB(0 To 0): ReDim sSrc(0 To 0): ReDim sVerIdx(0 To 0)
    ReDim sXD(0 To 0, 0 To 7)
    ReadTokjuk = 0

    Set ws = FindTokjuk()
    If ws Is Nothing Then
        AddErr "토적표 시트를 찾지 못했습니다."
        Exit Function
    End If

    v = ws.Range(ws.Cells(1, 1), ws.Cells(700, 101)).Value
    For r = 6 To 700
        If Len(Trim$(CStr(v(r, 2) & ""))) > 0 Then n = n + 1
    Next r
    If n = 0 Then
        AddErr "토적표에 측점이 없습니다 (B6:B700)."
        Exit Function
    End If

    ReDim sSta(0 To n - 1): ReDim sGH(0 To n - 1): ReDim sInv(0 To n - 1)
    ReDim sDia(0 To n - 1): ReDim sCode(0 To n - 1): ReDim sSlope(0 To n - 1)
    ReDim sExc(0 To n - 1): ReDim sWet(0 To n - 1): ReDim sAr(0 To n - 1): ReDim sAb(0 To n - 1)
    ReDim sAsp(0 To n - 1): ReDim sCon(0 To n - 1): ReDim sAsc(0 To n - 1): ReDim sBodo(0 To n - 1)
    ReDim sSand(0 To n - 1): ReDim sSub(0 To n - 1): ReDim sFrost(0 To n - 1)
    ReDim sB(0 To n - 1): ReDim sSrc(0 To n - 1): ReDim sVerIdx(0 To n - 1)
    ReDim sXD(0 To n - 1, 0 To 7)

    i = 0
    For r = 6 To 700
        If Len(Trim$(CStr(v(r, 2) & ""))) > 0 Then
            sSta(i) = Trim$(CStr(v(r, 2) & ""))
            sGH(i) = Dv(v(r, 5))
            sInv(i) = Dv(v(r, 7))
            sExc(i) = Dv(v(r, 9))
            sWet(i) = Dv(v(r, 10))
            sAr(i) = Dv(v(r, 11))
            sAb(i) = Dv(v(r, 12))
            sAsp(i) = Dv(v(r, 13))
            sCon(i) = Dv(v(r, 14))
            sAsc(i) = Dv(v(r, 15))
            sBodo(i) = Dv(v(r, 16))
            sSand(i) = Dv(v(r, 17))
            sSub(i) = Dv(v(r, 18))
            sFrost(i) = Dv(v(r, 19))
            sDia(i) = Dv(v(r, 96))
            sB(i) = Dv(v(r, 99))
            sCode(i) = Trim$(CStr(v(r, 100) & ""))
            sSlope(i) = pSlope
            sVerIdx(i) = r
            sSrc(i) = "TOKJUK"
            sXD(i, 0) = v(r, 94)      ' CP 라인명
            sXD(i, 1) = v(r, 95)      ' CQ 굴착장비
            sXD(i, 2) = v(r, 96)      ' CR 관경1
            sXD(i, 3) = v(r, 97)      ' CS 관경2
            sXD(i, 4) = v(r, 98)      ' CT 굴착심도
            sXD(i, 5) = v(r, 99)      ' CU 굴착폭
            sXD(i, 6) = v(r, 100)     ' CV 구간
            sXD(i, 7) = v(r, 101)     ' CW 관번호
            i = i + 1
        End If
    Next r

    sN = n
    ReadTokjuk = n
End Function

' SDB(H) : A열 = "S" 인 행이 단면 1개 (안내행 포함 2행/단면)
Private Function ReadSdbH() As Long
    Dim ws As Worksheet, v As Variant, r As Long, n As Long
    ReadSdbH = 0
    hN = 0
    Set ws = FindSdbH()
    If ws Is Nothing Then Exit Function

    v = ws.Range(ws.Cells(1, 1), ws.Cells(1400, 7)).Value
    For r = 2 To 1400
        If UCase$(Trim$(CStr(v(r, 1) & ""))) = "S" Then n = n + 1
    Next r
    If n = 0 Then Exit Function

    ReDim hSta(0 To n - 1): ReDim hGH(0 To n - 1): ReDim hInv(0 To n - 1)
    ReDim hDia(0 To n - 1): ReDim hCode(0 To n - 1): ReDim hSlope(0 To n - 1)
    n = 0
    For r = 2 To 1400
        If UCase$(Trim$(CStr(v(r, 1) & ""))) = "S" Then
            hSta(n) = Trim$(CStr(v(r, 2) & ""))
            hGH(n) = Dv(v(r, 3))
            hInv(n) = Dv(v(r, 4))
            hDia(n) = Dv(v(r, 5))
            hCode(n) = Trim$(CStr(v(r, 6) & ""))
            hSlope(n) = Dv(v(r, 7))
            If hSlope(n) <= 0# Then hSlope(n) = pSlope
            n = n + 1
        End If
    Next r
    hN = n
    ReadSdbH = n
End Function

' 측점 정규화 : "No.0+8.70" -> "0+8.70", "0+8.70(전)" -> "0+8.70", 0 -> "0"
Private Function NormSta(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    t = Replace(t, "NO.", ""): t = Replace(t, "No.", ""): t = Replace(t, "no.", "")
    t = Replace(t, vbTab, ""): t = Replace(t, " ", "")
    If Right$(t, 3) = "(전)" Or Right$(t, 3) = "(후)" Then t = Left$(t, Len(t) - 3)
    If Right$(t, 2) = "(전" Or Right$(t, 2) = "(후" Then t = Left$(t, Len(t) - 2)
    If Right$(t, 1) = "(" Then t = Left$(t, Len(t) - 1)
    If IsNumeric(t) Then t = Format$(CDbl(t), "0.####")
    NormSta = t
End Function

Private Function OccurIndex(ByRef arr() As String, ByVal upto As Long, ByVal key As String) As Long
    Dim i As Long, c As Long
    For i = 0 To upto - 1
        If NormSta(arr(i)) = key Then c = c + 1
    Next i
    OccurIndex = c
End Function

' SDB(H) 를 토적표에 매칭 (측점 + 등장순번) 하여 기하값만 덮어쓴다
Private Function MergeSections() As Long
    Dim i As Long, j As Long, n As Long, hit As Long, miss As Long
    Dim key As String
    Dim used() As Boolean
    Dim seen As Object
    Dim occ As Long

    MergeSections = 0
    If hN = 0 Then
        For i = 0 To sN - 1
            sSrc(i) = "TOKJUK"
        Next i
        AddWarn "SDB(H) 시트가 없어 토적표 값만 사용합니다."
        Exit Function
    End If

    Set seen = CreateObject("Scripting.Dictionary")
    ReDim used(0 To hN - 1)
    For i = 0 To sN - 1
        key = NormSta(sSta(i))
        occ = 0
        If seen.Exists(key) Then occ = seen(key)
        seen(key) = occ + 1

        hit = -1
        For j = 0 To hN - 1
            If Not used(j) Then
                If NormSta(hSta(j)) = key Then
                    If OccurIndex(hSta, j, key) = occ Then
                        hit = j
                        Exit For
                    End If
                End If
            End If
        Next j

        If hit >= 0 Then
            used(hit) = True
            sGH(i) = hGH(hit)
            sInv(i) = hInv(hit)
            sDia(i) = hDia(hit)
            If hCode(hit) <> "" Then sCode(i) = hCode(hit)
            sSlope(i) = hSlope(hit)
            sSrc(i) = "SDBH"
            n = n + 1
        Else
            miss = miss + 1
            sSrc(i) = "TOKJUK"
        End If
    Next i

    MergeSections = n
    AddWarn "SDB(H) 매칭 " & n & "건 / 토적표만 " & miss & "건"
End Function

Private Function LoadData() As Boolean
    LoadData = False
    LogInit
    ReadParams

    If ReadTokjuk() = 0 Then
        MsgBox "토적표를 읽지 못했습니다." & vbLf & vbLf & ErrText(), vbCritical, "중단"
        Exit Function
    End If

    Select Case UCase$(pSrcGeom)
    Case "SDBH"
        ReadSdbH
        MergeSections
    Case "AUTO"
        ReadSdbH
        If hN = sN And hN > 0 Then
            MergeSections
        Else
            AddWarn "SDB(H) " & hN & "행 / 토적표 " & sN & "행 이 달라 SDB(H)를 쓰지 않습니다."
        End If
    Case Else
        AddWarn "기하 소스 = 토적표"
    End Select
    LoadData = True
End Function

Private Function ErrText() As String
    ErrText = "Err " & Err.Number & " : " & Err.Description
End Function

'=====================================================================================
'  8. 포장 구성 해석   ("이름:두께:구분[:표시명];...")
'     구분 : CUT=포장(절토)층 / SUB=보조기층 / SAND=모래 / FROST=동상방지층
'=====================================================================================
Private Function NumS(ByVal s As String) As Double
    Dim t As String
    t = Trim$(s)
    t = Replace(t, ",", ".")
    If IsNumeric(t) Then NumS = CDbl(t) Else NumS = 0#
End Function

Private Function PaveDisp(ByVal nm As String) As String
    Select Case nm
    Case "표층":       PaveDisp = "표 층"
    Case "기층":       PaveDisp = "기 층"
    Case "보조기층":   PaveDisp = "보조기층"
    Case "보도블럭":   PaveDisp = "보도블럭"
    Case "모래":       PaveDisp = "모래부설"
    Case "모래부설":   PaveDisp = "모래부설"
    Case "동상방지층": PaveDisp = "동 상 방 지 층"
    Case Else:         PaveDisp = nm
    End Select
End Function

' 구간코드(OP02/OP04/SC01...) -> 포장코드("02"/"04"/"01"/"")
Private Function PaveCode(ByVal code As String) As String
    PaveCode = ""
    If Len(code) >= 4 Then PaveCode = Mid$(code, 3, 2)
End Function

Private Function HasPave(ByVal code As String) As Boolean
    Select Case PaveCode(code)
    Case "02":   HasPave = (Len(Trim$(pPave02)) > 0)
    Case "04":   HasPave = (Len(Trim$(pPave04)) > 0)
    Case Else:   HasPave = False
    End Select
End Function

Private Sub PaveParse(ByVal code As String)
    Dim s As String, f As Variant, a As Variant, i As Long
    On Error Resume Next
    Erase gLayName, gLayTh, gLayKind, gLayDisp
    On Error GoTo 0
    ReDim gLayName(0 To 15): ReDim gLayTh(0 To 15)
    ReDim gLayKind(0 To 15): ReDim gLayDisp(0 To 15)
    gNLay = 0

    Select Case PaveCode(code)
    Case "02": s = pPave02
    Case "04": s = pPave04
    Case Else: s = ""
    End Select
    If Len(Trim$(s)) = 0 Then Exit Sub

    f = Split(s, ";")
    For i = LBound(f) To UBound(f)
        If Len(Trim$(f(i))) > 0 Then
            a = Split(f(i), ":")
            If UBound(a) >= 2 Then
                gLayName(gNLay) = Trim$(CStr(a(0)))
                gLayTh(gNLay) = NumS(CStr(a(1)))
                gLayKind(gNLay) = UCase$(Trim$(CStr(a(2))))
                If UBound(a) >= 3 Then
                    If Len(Trim$(CStr(a(3)))) > 0 Then
                        gLayDisp(gNLay) = Trim$(CStr(a(3)))
                    Else
                        gLayDisp(gNLay) = PaveDisp(gLayName(gNLay))
                    End If
                Else
                    gLayDisp(gNLay) = PaveDisp(gLayName(gNLay))
                End If
                If gLayTh(gNLay) > 0# Then gNLay = gNLay + 1
            End If
        End If
    Next i
End Sub

'=====================================================================================
'  9. 단면 계산  (역설계 검증 완료 : 655단면 전수 일치)
'     · w(z) = b0 + 2*SLOPE*(H-z)      : 지반면에서 z 내려간 곳의 굴착폭
'     · 터파기 = (b0 + w(tv))/2 * (H-tv)
'     · 관상단 되메우기 = (w(zL)+w(zP))/2 * (zP-zL)
'     · 관주위 되메우기 = 터파기 - 비CUT층 - 관상단 - πD²/4
'=====================================================================================
Private Function MinD(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function
Private Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function

' VBA 의 Int() 는 음수에서 0 방향 절사(=trunc)가 아니라 진짜 floor 이지만,
' 부동소수 오차(예: 2.9999999999)를 막기 위해 EPS 보정 후 내림한다.
Private Function FloorD(ByVal x As Double) As Double
    Dim f As Double
    f = Int(x)
    If f > x Then f = f - 1#              ' 안전장치
    If x - f > 0.9999999 Then f = f + 1#  ' 2.9999999999 -> 3
    FloorD = f
End Function

Private Function BwOf(ByVal dia As Double) As Double
    If Abs(dia - 80#) < 0.5 Then
        BwOf = pBw80
    ElseIf Abs(dia - 100#) < 0.5 Then
        BwOf = pBw100
    ElseIf Abs(dia - 160#) < 0.5 Then
        BwOf = pBw160
    Else
        BwOf = pBwOther
    End If
End Function

' 지반면에서 z 만큼 내려간 곳의 굴착폭
Private Function Wz(ByVal z As Double) As Double
    Wz = gb0 + 2# * gSlp * (gH - z)
End Function

Private Sub SecCalc(ByVal k As Long)
    Dim D As Double, H As Double, b0 As Double
    Dim tv As Double, zL As Double, z As Double, i As Long
    Dim zt As Double, zb As Double, zt2 As Double, zb2 As Double
    Dim aLay As Double, sumLay As Double
    Dim zP As Double, ab As Double, ar As Double, exc As Double
    Dim wt As Double, slp As Double, yr As Double
    Dim pc As String

    gCurK = k
    PaveParse sCode(k)
    pc = PaveCode(sCode(k))

    D = sDia(k) / 1000#
    H = sGH(k) - sInv(k)
    slp = sSlope(k)
    If slp <= 0# Then slp = pSlope
    gSlp = slp
    b0 = D + 2# * BwOf(sDia(k))

    tv = 0#: zL = 0#
    For i = 0 To gNLay - 1
        zL = zL + gLayTh(i)
        If gLayKind(i) = "CUT" Then tv = tv + gLayTh(i)
    Next i

    gD = D: gH = H: gb0 = b0: gtv = tv
    ghb = b0 / 2#

    gExc = 0#: gWet = 0#: gAr = 0#: gAb = 0#
    gSand = 0#: gSub = 0#: gFrost = 0#
    gAsp = 0#: gCon = 0#: gAsc = 0#: gBodo = 0#
    gB = 0#: ght = ghb: gzL = 0#: gzP = 0#
    gzRfil = 0#: gRfilHalf = 0#
    gTopik = Rnd2(H - D)
    gHasExc = False

    If UCase$(Left$(sCode(k), 2)) = "OP" And H > 0.0000001 Then
        gHasExc = True

        wt = Wz(tv)                    ' 굴착 상단폭 w(tv)
        ght = wt / 2#
        gB = Rnd2(wt)
        exc = (b0 + wt) / 2# * (H - tv)
        gExc = Rnd2(exc)

        z = 0#: sumLay = 0#
        For i = 0 To gNLay - 1
            zt = z: zb = z + gLayTh(i): z = zb
            If gLayKind(i) <> "CUT" Then
                zt2 = MinD(MaxD(zt, tv), H)
                zb2 = MinD(MaxD(zb, tv), H)
                If zb2 > zt2 + 0.0000001 Then
                    aLay = (Wz(zt2) + Wz(zb2)) / 2# * (zb2 - zt2)
                    sumLay = sumLay + aLay
                    Select Case gLayKind(i)
                    Case "SAND":  gSand = gSand + aLay
                    Case "SUB":   gSub = gSub + aLay
                    Case "FROST": gFrost = gFrost + aLay
                    End Select
                End If
            End If
        Next i

        zL = z
        If gNLay > 0 Then zL = MinD(zL, H) Else zL = 0#
        zL = MaxD(zL, tv)
        zP = H - D

        If zP > zL + 0.0000001 Then
            ab = (Wz(zL) + Wz(zP)) / 2# * (zP - zL)
        Else
            ab = 0#
        End If

        ar = exc - sumLay - ab - PI * D * D / 4#

        gzL = zL
        gzP = zP
        gAb = Rnd2(ab)
        If ar >= 0# Or Not pNegZero Then
            gAr = Rnd2(ar)
        Else
            gAr = 0#
        End If
        gSand = Rnd2(gSand): gSub = Rnd2(gSub): gFrost = Rnd2(gFrost)

        ' 원본 규칙 : ASP포장(OP02) / 보도블럭포장(OP04) = 굴착상단폭
        If pc = "02" Then gAsp = gB
        If pc = "04" Then gBodo = gB
    End If

    ' ---- 관저 기준선 : y = max(zP,zL) ----
    yr = MaxD(gzP, gzL)
    If pRfilClamp Then yr = MinD(yr, H)
    If yr < 0# Then yr = 0#
    gzRfil = yr
    If gHasExc Then gRfilHalf = Wz(yr) / 2# Else gRfilHalf = 0#
End Sub

'=====================================================================================
'  12. 해치 (관주위 되메우기)  - 사다리꼴(바닥 b0/2 -> 상단 w(y)/2) - 관원
'      원본과 동일: 외곽=선 4개, 내곽=원 1개, ANSI31, 축척 0.02
'=====================================================================================
Private Sub HatchBackfill(ByVal cx As Double, ByVal cy As Double, ByVal circ As Object)
    Dim tmp As Object, h As Object, arr(0 To 0) As Object
    Dim y As Double, yh As Double
    Dim pts As String

    y = gzRfil
    yh = gRfilHalf
    If y <= TINY Or yh <= TINY Then Exit Sub

    ' 임시 폴리라인(닫힌 사다리꼴) -> 해치 외곽 -> 삭제
    pts = Pt(cx - ghb, cy - gH) & ";" & Pt(cx - yh, cy - y) & ";" & _
          Pt(cx + yh, cy - y) & ";" & Pt(cx + ghb, cy - gH) & ";" & Pt(cx - ghb, cy - gH)
    Set tmp = xcPoly(pts, pLayPatt, False)
    If tmp Is Nothing Then Exit Sub

    On Error Resume Next
    Set h = ms.AddHatch(0, pHatchPat, False)      ' 0 = 미리정의 패턴
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        tmp.Delete
        Exit Sub
    End If

    Set arr(0) = tmp
    h.AppendOuterLoop arr
    If Not circ Is Nothing Then
        Set arr(0) = circ
        h.AppendInnerLoop arr
    End If
    h.PatternScale = pHatchScale
    h.Evaluate
    h.Layer = pLayPatt
    If Err.Number <> 0 Then Err.Clear

    tmp.Delete                                    ' 경계는 비연관(assoc=False)이므로 유지됨
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

'=====================================================================================
'  13. 표 블록 (관로06) - 15개 속성 + Tojuk XDATA
'=====================================================================================
Private Sub DrawTableBlock(ByVal cx As Double, ByVal cy As Double, ByVal k As Long)
    Dim blk As Object, v As Variant, i As Long
    Dim s As String
    Dim xd(0 To 7) As String
    Dim xt(0 To 7) As Integer

    Set blk = Nothing
    On Error Resume Next
    Set blk = ms.InsertBlock(xcP3(cx + pTblDX, cy), pBlkTable, pBlkScale, pBlkScale, pBlkScale, 0#)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        AddErr "표 블록 삽입 실패 (k=" & k & ") : " & pBlkTable
        Exit Sub
    End If
    blk.Layer = pLayTabl
    On Error GoTo 0

    ' ---- 15개 속성 ----
    On Error Resume Next
    If blk.HasAttributes Then
        v = blk.GetAttributes
        For i = LBound(v) To UBound(v)
            Select Case v(i).TagString
            Case "측점":               s = sSta(k)
            Case "지반고":             s = F2(sGH(k))
            Case "관저고":             s = F2(sInv(k))
            Case "토피":               s = F2(gTopik)
            Case "터파기(육상)합계":   s = F2(gExc)
            Case "터파기(용수)합계":   s = F2(gWet)
            Case "되메우기(관주위)":   s = F2(gAr)
            Case "되메우기(관상단)":   s = F2(gAb)
            Case "ASP포장":            s = F2(gAsp)
            Case "CONC포장":           s = F2(gCon)
            Case "ASC포장":            s = F2(gAsc)
            Case "보도블럭포장":       s = F2(gBodo)
            Case "되메우기(모래)":     s = F2(gSand)
            Case "보조기층":           s = F2(gSub)
            Case "동상방지층":         s = F2(gFrost)
            Case Else:                 s = v(i).TextString
            End Select
            v(i).TextString = s
        Next i
    End If
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    ' ---- Tojuk XDATA (8개 문자열) ----
    If pWriteXdata And Len(pXdataApp) > 0 Then
        xd(0) = FI(sXD(k, 0))          ' 라인명
        xd(1) = FI(sXD(k, 1))          ' 굴착장비
        xd(2) = FI(sXD(k, 2))          ' 관경1
        xd(3) = FI(sXD(k, 3))          ' 관경2
        xd(4) = FI(sXD(k, 4))          ' 굴착심도
        xd(5) = FI(sXD(k, 5))          ' 굴착폭
        xd(6) = FI(sXD(k, 6))          ' 구간
        xd(7) = FI(sXD(k, 7))          ' 관번호
        For i = 0 To 7
            xt(i) = 1000                ' 1000 = 문자열
        Next i
        On Error Resume Next
        blk.SetXData pXdataApp, xt, xd
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
    End If
End Sub

'=====================================================================================
'  14. 고도축  - 하단 = floor(INV) - EL_BOT_OFF, 상단 = 하단 + EL_TOP_MIN + floor(H)
'=====================================================================================
Private Sub DrawElAxis(ByVal cx As Double, ByVal cy As Double, ByVal k As Long)
    Dim bot As Double, top As Double, z As Double
    Dim yb As Double, yt As Double
    Dim xax As Double
    Dim gh As Double, h As Double

    gh = sGH(k)
    h = sGH(k) - sInv(k)

    ' 원본 규칙 : 하단 = floor(관저고) - EL_BOT_OFF,  상단 = 하단 + EL_TOP_MIN + floor(H)
    bot = FloorD(sInv(k)) - pElBotOff
    top = bot + pElTopMin + FloorD(h)

    xax = cx - 4.5
    yb = cy + (bot - gh)
    yt = cy + (top - gh)

    If yt <= yb Then yt = yb + 1#

    ' 축선
    xcLine xax, yb, xax, yt, pLayDiml

    ' 눈금 (1 m 간격)
    z = bot
    Do While z <= top + 0.0000001
        xcLine xax, cy + (z - gh), xax - 0.4, cy + (z - gh), pLayDiml
        z = z + 1#
    Loop

    ' 상·하 단 2개만 숫자 표기 (우측 정렬)
    xcText FmtInt(bot), cx - 5#, yb, pTxtEl, pLayText, 0#, AL_MRGT
    xcText FmtInt(top), cx - 5#, yt, pTxtEl, pLayText, 0#, AL_MRGT
End Sub

' 정수 표기용 (고도축 눈금 숫자)
Private Function FmtInt(ByVal v As Double) As String
    Dim n As Long
    n = CLng(FloorD(v + 0.0000001))
    FmtInt = CStr(n)
End Function

'=====================================================================================
'  15. 상세치수 (첫단면만, 28개 엔티티)
'      A) B= 폭 치수 (상단)      B) H= 높이 치수 (좌측)
'      C) 저부 관경/기초폭       D) 외측 폭 치수 (우측)
'=====================================================================================
Private Sub DrawDetailDim(ByVal cx As Double, ByVal cy As Double)
    Dim x As Double, y As Double, d As Double
    Dim bw As Double

    ' ---- A) B= 폭 치수 : y=+1.9, 반폭 = ght ----
    If True Then
        xcLine cx - ght, cy + 1.9, cx + ght, cy + 1.9, pLayLead
        xcLine cx - ght, cy + 1.9, cx - ght, cy + 1.5, pLayLead
        xcLine cx + ght, cy + 1.9, cx + ght, cy + 1.5, pLayLead
        Dot cx - ght, cy + 1.9
        Dot cx + ght, cy + 1.9
        xcText "B=" & Format$(gB, "0.00") & "m", cx, cy + 2.1, pTxtDim, pLayLeat, 0#, AL_MCTR
    End If

    ' ---- B) H= 높이 치수 : x=-1.9 ----
    xcLine cx - 1.9, cy, cx - 1.9, cy - gH, pLayLead
    xcLine cx - 1.9, cy, cx - 1.5, cy, pLayLead
    xcLine cx - 1.9, cy - gH, cx - 1.5, cy - gH, pLayLead
    Dot cx - 1.9, cy
    Dot cx - 1.9, cy - gH
    xcText "H=" & Format$(gH, "0.00") & "m", cx - 2.1, cy - gH / 2#, pTxtDim, pLayLeat, 90#, AL_MCTR

    ' ---- C) 저부 치수 ----
    If pDetailBottom Then
        y = cy - gH - 1.0
        d = cy - gH - 1.4
        bw = BwOf(sDia(gCurK))

        xcLine cx - ghb, d, cx + ghb, d, pLayLead
        xcLine cx - ghb, y, cx - ghb, d, pLayLead
        xcLine cx + ghb, y, cx + ghb, d, pLayLead
        If bw > TINY Then
            xcLine cx - ghb + bw, y, cx - ghb + bw, d, pLayLead
            xcLine cx + ghb - bw, y, cx + ghb - bw, d, pLayLead
            Dot cx - ghb + bw, d
            Dot cx + ghb - bw, d
        End If
        Dot cx - ghb, d
        Dot cx + ghb, d

        x = ghb + 0.1
        If bw <= TINY Then x = ghb + 0.2
        xcText Format$(bw * 1000#, "0"), cx - x, cy - gH - 1.6, pTxtPave, pLayLeat, 0#, AL_MCTR
        xcText Format$(bw * 1000#, "0"), cx + x, cy - gH - 1.6, pTxtPave, pLayLeat, 0#, AL_MCTR
        xcText Format$(gD * 1000#, "0"), cx, cy - gH - 1.2, pTxtPave, pLayLeat, 0#, AL_MCTR
    End If

    ' ---- D) 외측 폭 치수 (우측) ----
    y = cy - gzRfil
    d = cy - gH
    x = ght + 0.45
    If x < 0.88 Then x = 0.88

    xcLine cx + x, y, cx + x, d, pLayLead
    xcLine cx + x - 0.4, y, cx + x, y, pLayLead
    xcLine cx + x - 0.4, d, cx + x, d, pLayLead
    Dot cx + x, y
    Dot cx + x, d

    xcText "0", cx + x + 0.2, y + 0.2, pTxtPave, pLayLeat, 90#, AL_MCTR
    xcText Format$(gD * 1000#, "0"), cx + x - 0.2, (y + d) / 2#, pTxtPave, pLayLeat, 90#, AL_MCTR
    xcText "0", cx + x + 0.2, d - 0.2, pTxtPave, pLayLeat, 90#, AL_MCTR
End Sub

Private Sub Dot(ByVal x As Double, ByVal y As Double)
    On Error Resume Next
    ms.InsertBlock xcP3(x, y), pBlkDot, pDotScale, pDotScale, pDotScale, 0#
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

'=====================================================================================
'  16. 포장 리더 (첫단면 / 포장변경)
'      · 세로선 x=+3.3 : 최하층 중앙 ~ (2.5 + 최하층중앙)
'      · 각 층 중앙높이에 점
'      · 상단에 0.3 간격 눈금 3개 + "표 층(T=50)" 형식 문자
'=====================================================================================
Private Sub DrawPaveLeader(ByVal cx As Double, ByVal cy As Double)
    Dim i As Long, z As Double, midY As Double, lastMid As Double
    Dim xv As Double, top As Double, yt As Double

    If gNLay = 0 Then Exit Sub
    xv = cx + 3.3

    z = 0#: lastMid = 0#
    For i = 0 To gNLay - 1
        z = z + gLayTh(i)
        midY = cy - (z - gLayTh(i) / 2#)
        lastMid = midY
        Dot xv, midY
    Next i

    top = lastMid + 2.5
    xcLine xv, lastMid, xv, top, pLayLeal

    For i = 0 To gNLay - 1
        yt = top - 0.3 * i
        xcLine xv, yt, xv + 0.5, yt, pLayLeal
        xcText gLayDisp(i) & "(T=" & Format$(gLayTh(i) * 1000#, "0") & ")", _
                xv + 0.5, yt - 0.1, pTxtPave, pLayLeat, 0#, AL_MLFT
    Next i
End Sub

'=====================================================================================
'  17. 도곽 · 표제란
'=====================================================================================
Private Sub DrawSheetBorder(ByVal n As Long)
    Dim ox As Double, oy As Double, i As Long
    Dim blk As Object
    Dim p As Variant

    ox = gBaseX + (n - 1) * pSheetW
    oy = gBaseY

    ' 1) XREF 로 붙이기 (BORDER_XREF 가 있을 때)
    If Len(Trim$(pBorderXref)) > 0 Then
        If XrefBorder(ox, oy) Then Exit Sub
    End If

    ' 2) 블록 삽입
    On Error Resume Next
    Set blk = ms.InsertBlock(xcP3(ox, oy), pBlkBorder, pBlkScale, pBlkScale, pBlkScale, 0#)
    If Err.Number = 0 And Not blk Is Nothing Then
        blk.Layer = pLayBorder
        On Error GoTo 0
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0

    ' 3) 외곽선 직접 작도
    If pBorderFallback Then
        xcPoly Pt(ox, oy) & ";" & Pt(ox + pBorderW, oy) & ";" & _
                Pt(ox + pBorderW, oy + pBorderH) & ";" & Pt(ox, oy + pBorderH), _
                pLayBorder, True
        AddWarn "도곽 블록(" & pBlkBorder & ")이 없어 외곽선으로 대체했습니다 (시트 " & n & ")"
    End If
End Sub

Private Function XrefBorder(ByVal ox As Double, ByVal oy As Double) As Boolean
    Dim blk As Object, xr As Object
    XrefBorder = False
    On Error Resume Next
    Set blk = doc.Blocks.Add(xcP3(0#, 0#), "TMP_BORDER_XREF")
    If Err.Number <> 0 Then Err.Clear: Exit Function
    Set xr = blk.AttachExternalReference(pBorderXref, "XREF_BORDER_TPL", xcP3(0#, 0#), 1#, 1#, 1#, 0#, False)
    If Err.Number <> 0 Then
        Err.Clear
        doc.Blocks.Item("TMP_BORDER_XREF").Delete
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    ' 템플릿 블록을 각 시트 위치에 삽입
    Dim b As Object
    Set b = ms.InsertBlock(xcP3(ox, oy), "TMP_BORDER_XREF", pBlkScale, pBlkScale, pBlkScale, 0#)
    If Err.Number = 0 And Not b Is Nothing Then
        b.Layer = pLayBorder
        XrefBorder = True
    End If
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function

Private Sub DrawSheetTitle(ByVal n As Long)
    Dim ox As Double, oy As Double
    Dim blk As Object, v As Variant, i As Long

    ox = gBaseX + (n - 1) * pSheetW
    oy = gBaseY

    Set blk = Nothing
    On Error Resume Next
    Set blk = ms.InsertBlock(xcP3(ox, oy), pBlkTitle, pBlkScale, pBlkScale, pBlkScale, 0#)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        AddErr "표제란 블록 삽입 실패 (시트 " & n & ") : " & pBlkTitle
        Exit Sub
    End If
    blk.Layer = pLayTitle

    If blk.HasAttributes Then
        v = blk.GetAttributes
        For i = LBound(v) To UBound(v)
            Select Case v(i).TagString
            Case "도면타이틀": v(i).TextString = TitleText(n)
            Case "도면명":     v(i).TextString = TitleText(n)
            Case "축척H":      v(i).TextString = pScaleH
            Case "축척V":      v(i).TextString = ""
            Case "축척A1":     v(i).TextString = pScaleA1
            Case "SNO1":       v(i).TextString = ""
            Case "SNO2":       v(i).TextString = ""
            Case "도면번호":   v(i).TextString = pDwgPrefix & Format$(pDwgStart + n - 1, "0")
            End Select
        Next i
    End If
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Private Function TitleText(ByVal n As Long) As String
    TitleText = Replace(pTitleFmt, "{n}", Format$(n, "0"))
End Function

'=====================================================================================
'  18. 기준점 선택
'=====================================================================================
Private Function xcPickBase() As Boolean
    Dim p As Variant
    Dim msg As String
    xcPickBase = False

    On Error Resume Next
    acad.Visible = True
    On Error GoTo 0

    On Error Resume Next
    AppActivate acad.Caption
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    msg = vbLf & "횡단면도 시작점을 클릭하세요  [ESC = 취소]" & vbLf & _
          "  → 이 점이 1번 시트의 좌하 모서리(도곽 원점)가 됩니다." & vbLf

    On Error Resume Next
    p = doc.Utility.GetPoint(, msg)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        MsgBox "기준점 선택을 취소했습니다.", vbInformation, "취소"
        Exit Function
    End If
    On Error GoTo 0

    gBaseX = CDbl(p(0))
    gBaseY = CDbl(p(1))
    xcPickBase = True
End Function

'=====================================================================================
'  19. 검증 시트 (검토_검증)
'=====================================================================================
Private Sub xcWriteCheck()
    Dim ws As Worksheet, r As Long, i As Long
    Dim nBad As Long

    On Error Resume Next
    Application.DisplayAlerts = False
    Set ws = ThisWorkbook.Worksheets(VSHEET)
    If Not ws Is Nothing Then ws.Delete
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Count))
    ws.Name = VSHEET
    Application.DisplayAlerts = True

    ws.Range("A1").Value = "횡단면도 데이터 검증"
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 13

    ws.Range("A3:G3").Value = Array("No", "측점", "구간", "지반고", "관저고", "관경", "기하소스", "H(토피)", "B(굴착폭)", "터파기", "되메우기(관주위)", "비고")
    ws.Range("A3:L3").Font.Bold = True
    ws.Range("A3:L3").Interior.Color = RGB(204, 255, 204)

    r = 4
    For i = 0 To sN - 1
        SecCalc i
        ws.Cells(r, 1).Value = i + 1
        ws.Cells(r, 2).Value = "'" & sSta(i)
        ws.Cells(r, 3).Value = sCode(i)
        ws.Cells(r, 4).Value = sGH(i)
        ws.Cells(r, 5).Value = sInv(i)
        ws.Cells(r, 6).Value = sDia(i)
        ws.Cells(r, 7).Value = sSrc(i)
        ws.Cells(r, 8).Value = Rnd2(gH)
        ws.Cells(r, 9).Value = gB
        ws.Cells(r, 10).Value = gExc
        ws.Cells(r, 11).Value = gAr

        ' 토적표 값과 비교
        If sSrc(i) = "SDBH" Then
            If Abs(gExc - sExc(i)) > 0.005 And sExc(i) > 0 Then
                ws.Cells(r, 12).Value = "터파기 차이 " & Format$(gExc - sExc(i), "0.00")
                nBad = nBad + 1
            End If
            If Abs(gAr - sAr(i)) > 0.005 Then
                ws.Cells(r, 12).Value = ws.Cells(r, 12).Value & " 관주위 차이 " & Format$(gAr - sAr(i), "0.00")
                nBad = nBad + 1
            End If
        End If
        r = r + 1
    Next i

    r = r + 1
    ws.Cells(r, 1).Value = "합계"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 2).Value = sN & "단면"
    ws.Cells(r, 3).Value = "불일치 " & nBad & "건"
    If nBad > 0 Then
        ws.Cells(r, 3).Font.Color = RGB(192, 0, 0)
    Else
        ws.Cells(r, 3).Font.Color = RGB(0, 128, 0)
    End If

    r = r + 2
    ws.Cells(r, 1).Value = "경고/오류"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1
    If Not gWarn Is Nothing Then
        For i = 1 To gWarn.Count
            ws.Cells(r, 1).Value = "·"
            ws.Cells(r, 2).Value = gWarn(i)
            r = r + 1
        Next i
    End If
    If Not gErr Is Nothing Then
        For i = 1 To gErr.Count
            ws.Cells(r, 1).Value = "!"
            ws.Cells(r, 2).Value = gErr(i)
            ws.Cells(r, 2).Font.Color = RGB(192, 0, 0)
            r = r + 1
        Next i
    End If

    With ws
        .Columns("A").ColumnWidth = 6
        .Columns("B").ColumnWidth = 16
        .Columns("C:G").ColumnWidth = 12
        .Columns("H:K").ColumnWidth = 14
        .Columns("L").ColumnWidth = 40
        .Range(.Cells(3, 1), .Cells(r, 12)).VerticalAlignment = xlTop
    End With
End Sub

'=====================================================================================
'  20. 진입점
'=====================================================================================
' (1) 데이터만 검증
Public Sub 횡단_1단계_검증()
    If LoadData() Then
        xcWriteCheck
        MsgBox "검증 완료" & vbLf & vbLf & _
               "· 단면 수 : " & sN & vbLf & _
               "· 결과 시트 : " & VSHEET, vbInformation, "1단계 검증"
    End If
End Sub

' (2) 전체 작도
Public Sub 횡단_작도()
    Dim i As Long, nSheet As Long, t0 As Single
    Dim msg As String

    If Not LoadData() Then Exit Sub

    nSheet = (sN + pNCol * pNRow - 1) \ (pNCol * pNRow)
    msg = "다음 내용으로 작도합니다." & vbLf & vbLf & _
          "· 단면 수 : " & sN & vbLf & _
          "· 도면 장수 : " & nSheet & " 장  (" & pNCol & "열 x " & pNRow & "행 = " & _
          (pNCol * pNRow) & "단면/장)" & vbLf & _
          "· 기하 소스 : " & pSrcGeom & vbLf & _
          "· 표 블록 : " & pBlkTable & "  (축척 " & pBlkScale & ")" & vbLf & vbLf & _
          "계속할까요?"
    If MsgBox(msg, vbYesNo + vbQuestion, "횡단면도 작도") <> vbYes Then Exit Sub

    If Not xcConnectCAD() Then Exit Sub

    If pPickBase Then
        If Not xcPickBase() Then Exit Sub
    Else
        gBaseX = pBaseX: gBaseY = pBaseY
    End If

    t0 = Timer
    SetupLayers
    CreateTableBlock
    CreateTitleBlock
    CreateDotBlock

    If pDbgSkipDraw Then
        xcWriteCheck
        MsgBox "DBG_SKIP_DRAW = TRUE 이므로 작도하지 않았습니다.", vbInformation, "테스트"
        Exit Sub
    End If

    ' ---- 시트 도곽 + 표제란 ----
    For i = 1 To nSheet
        DrawSheetBorder i
        DrawSheetTitle i
    Next i

    ' ---- 단면 ----
    For i = 0 To sN - 1
        DrawSection i
        If i Mod 50 = 0 Then
            On Error Resume Next
            Application.StatusBar = "작도중 " & (i + 1) & " / " & sN
            On Error GoTo 0
        End If
    Next i
    On Error Resume Next
    Application.StatusBar = False
    On Error GoTo 0

    If pVerify Then xcWriteCheck

    On Error Resume Next
    doc.Regen 1
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    MsgBox "작도 완료" & vbLf & vbLf & _
           "· 단면 " & sN & "개 / " & nSheet & "장" & vbLf & _
           "· 소요 " & Format$(Timer - t0, "0.0") & " 초" & vbLf & _
           "· 캐드 : " & cadAppName, vbInformation, "완료"
End Sub

' (3) 표 블록 속성만 갱신 (도면 재작성 없이)
Public Sub 횡단_블록갱신()
    Dim i As Long, j As Long, k As Long
    Dim e As Object, v As Variant
    Dim cx As Double, cy As Double
    Dim nUpd As Long, nSkip As Long
    Dim s As String

    If Not LoadData() Then Exit Sub
    If Not xcConnectCAD() Then Exit Sub

    For Each e In ms
        If e.EntityName = "AcDbBlockReference" Then
            If UCase$(e.Name) = UCase$(pBlkTable) Then
                ' 삽입점 -> 셀 역산
                cx = e.InsertionPoint(0) - pTblDX
                cy = e.InsertionPoint(1)
                k = CellToIndex(cx, cy)
                If k < 0 Then
                    nSkip = nSkip + 1
                Else
                    SecCalc k
                    If e.HasAttributes Then
                        v = e.GetAttributes
                        For j = LBound(v) To UBound(v)
                            Select Case v(j).TagString
                            Case "지반고":             s = F2(sGH(k))
                            Case "관저고":             s = F2(sInv(k))
                            Case "토피":               s = F2(gTopik)
                            Case "터파기(육상)합계":   s = F2(gExc)
                            Case "터파기(용수)합계":   s = F2(gWet)
                            Case "되메우기(관주위)":   s = F2(gAr)
                            Case "되메우기(관상단)":   s = F2(gAb)
                            Case "ASP포장":            s = F2(gAsp)
                            Case "CONC포장":           s = F2(gCon)
                            Case "ASC포장":            s = F2(gAsc)
                            Case "보도블럭포장":       s = F2(gBodo)
                            Case "되메우기(모래)":     s = F2(gSand)
                            Case "보조기층":           s = F2(gSub)
                            Case "동상방지층":         s = F2(gFrost)
                            Case Else:                 s = ""
                            End Select
                            If Len(s) > 0 Then v(j).TextString = s
                        Next j
                        nUpd = nUpd + 1
                    End If
                End If
            End If
        End If
    Next e

    MsgBox "표 블록 갱신 완료" & vbLf & vbLf & _
           "· 갱신 : " & nUpd & " 개" & vbLf & _
           "· 건너뜀 : " & nSkip & " 개", vbInformation, "블록갱신"
End Sub

' 삽입점 -> 단면 인덱스 (없으면 -1)
Private Function CellToIndex(ByVal cx As Double, ByVal cy As Double) As Long
    Dim sh As Long, col As Long, row As Long, k As Long
    Dim rx As Double, ry As Double

    CellToIndex = -1
    If pPitchX <= 0# Or pPitchY <= 0# Then Exit Function

    sh = Int((cx - gBaseX) / pSheetW + 0.0001)
    If sh < 0 Then sh = 0
    rx = cx - (gBaseX + sh * pSheetW) - pOffX
    ry = cy - gBaseY - pOffY
    col = Int(rx / pPitchX + 0.5)
    row = Int(ry / pPitchY + 0.5)
    If col < 0 Or col >= pNCol Then Exit Function
    If row < 0 Or row >= pNRow Then Exit Function

    k = sh * (pNCol * pNRow) + col * pNRow + row
    If k < 0 Or k > sN - 1 Then Exit Function
    CellToIndex = k
End Function

' (4) 도곽 · 블록 상태 점검
Public Sub 횡단_도곽확인()
    Dim e As Object
    Dim nTab As Long, nBor As Long, nTit As Long, nDot As Long
    Dim nXref As Long
    Dim s As String
    Dim b As Object

    If Not xcConnectCAD() Then Exit Sub

    For Each e In ms
        If e.EntityName = "AcDbBlockReference" Then
            If UCase$(e.Name) = UCase$(pBlkTable) Then nTab = nTab + 1
            If UCase$(e.Name) = UCase$(pBlkDot) Then nDot = nDot + 1
            If UCase$(e.Name) = UCase$(pBlkBorder) Then nBor = nBor + 1
            If UCase$(e.Name) = UCase$(pBlkTitle) Then nTit = nTit + 1
            On Error Resume Next
            If e.IsXRef Then nXref = nXref + 1
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
        End If
    Next e

    s = "도면 점검 결과" & vbLf & vbLf & _
        "· 캐드 : " & cadAppName & vbLf & _
        "· 표 블록(" & pBlkTable & ") : " & nTab & " 개" & vbLf & _
        "· 도곽(" & pBlkBorder & ") : " & nBor & " 개" & vbLf & _
        "· 표제란(" & pBlkTitle & ") : " & nTit & " 개" & vbLf & _
        "· 보조점(" & pBlkDot & ") : " & nDot & " 개" & vbLf & _
        "· XREF : " & nXref & " 개" & vbLf & vbLf

    s = s & "블록 정의" & vbLf
    s = s & "  " & pBlkTable & " : " & YesNo(BlockExists(pBlkTable)) & vbLf
    s = s & "  " & pBlkBorder & " : " & YesNo(BlockExists(pBlkBorder)) & vbLf
    s = s & "  " & pBlkTitle & " : " & YesNo(BlockExists(pBlkTitle)) & vbLf
    s = s & "  " & pBlkDot & " : " & YesNo(BlockExists(pBlkDot)) & vbLf & vbLf
    s = s & "※ 없는 블록은 [횡단_작도] 실행 시 자동 생성됩니다."

    MsgBox s, vbInformation, "도곽확인"
End Sub

Private Function YesNo(ByVal b As Boolean) As String
    If b Then YesNo = "있음" Else YesNo = "없음"
End Function

'=====================================================================================
'  21. 설정 초기화 (P 시트가 없을 때 자동 호출)
'=====================================================================================
Private Sub EnsureParams()
    Dim ws As Worksheet
    Set ws = SheetByName(PSHEET)
    If ws Is Nothing Then
        InitParams
    Else
        ReadParams
    End If
End Sub
'=====================================================================================
'  블록 자체 생성 (원본 DWG 에 블록이 없어도 동작하도록)
'  ------------------------------------------------------------------------------
'  좌표는 원본 DXF 블록정의에서 그대로 추출한 값이다. 임의로 고치지 말 것.
'=====================================================================================
Private Function BlockExists(ByVal nm As String) As Boolean
    Dim b As Object
    BlockExists = False
    If Len(nm) = 0 Then Exit Function
    On Error Resume Next
    Set b = doc.Blocks.Item(nm)
    If Err.Number <> 0 Then Err.Clear: Exit Function
    On Error GoTo 0
    BlockExists = Not (b Is Nothing)
End Function

' 관로06 : 70.0 x 35.2 표 블록  (LINE 15, TEXT 41, ATTDEF 15)
Public Sub CreateTableBlock()
    Dim blk As Object, e As Object, at As Object
    If BlockExists(pBlkTable) Then Exit Sub
    On Error Resume Next
    Set blk = doc.Blocks.Add(xcP3(0#, 0#), pBlkTable)
    If Err.Number <> 0 Then Err.Clear: Exit Sub
    On Error GoTo 0

    ' --- 외곽 (닫힌 폴리라인) ---
    Dim a0(0 To 7) As Double
    a0(0) = 0: a0(1) = 0
    a0(2) = 70: a0(3) = 0
    a0(4) = 70: a0(5) = -35.2
    a0(6) = 0: a0(7) = -35.2
    Set e = blk.AddLightWeightPolyline(a0)
    e.Closed = True

    ' --- 내부 구분선 15개 ---
    blk.AddLine xcP3(35, -4.4), xcP3(35, -35.2)
    blk.AddLine xcP3(57, -4.4), xcP3(57, -35.2)
    blk.AddLine xcP3(46, -4.4), xcP3(46, -22)
    blk.AddLine xcP3(11, -17.6), xcP3(11, -35.2)
    blk.AddLine xcP3(0, -17.6), xcP3(35, -17.6)
    blk.AddLine xcP3(11, -22), xcP3(70, -22)
    blk.AddLine xcP3(0, -8.8), xcP3(35, -8.8)
    blk.AddLine xcP3(0, -26.4), xcP3(70, -26.4)
    blk.AddLine xcP3(22, -4.4), xcP3(22, -35.2)
    blk.AddLine xcP3(0, -4.4), xcP3(70, -4.4)
    blk.AddLine xcP3(46, -17.6), xcP3(70, -17.6)
    blk.AddLine xcP3(46, -13.2), xcP3(70, -13.2)
    blk.AddLine xcP3(46, -8.8), xcP3(70, -8.8)
    blk.AddLine xcP3(11, -30.8), xcP3(70, -30.8)
    blk.AddLine xcP3(0, -13.2), xcP3(35, -13.2)

    ' --- 고정 문자 41개 ---
    Set e = blk.AddText("2", xcP3(68.89244, -28.401175), 1)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(69.16478, -27.901175)
    Set e = blk.AddText("m", xcP3(67.096308, -29.7), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(67.695457, -28.6)
    Set e = blk.AddText("m", xcP3(67.096308, -20.9), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(67.695457, -19.8)
    Set e = blk.AddText("2", xcP3(68.89244, -24.001175), 1)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(69.16478, -23.501175)
    Set e = blk.AddText("m", xcP3(67.096308, -25.3), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(67.695457, -24.2)
    Set e = blk.AddText("m", xcP3(67.062365, -12.1), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(67.661514, -11)
    Set e = blk.AddText("m", xcP3(67.096308, -7.7), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(67.695457, -6.6)
    Set e = blk.AddText("보  조  기  층", xcP3(37.611915, -29.7), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(46, -28.6)
    Set e = blk.AddText("2", xcP3(33.89244, -32.801175), 1)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(34.16478, -32.301175)
    Set e = blk.AddText("2", xcP3(33.89244, -28.401175), 1)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(34.16478, -27.901175)
    Set e = blk.AddText("모  래  부  설", xcP3(37.611915, -25.3), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(46, -24.2)
    Set e = blk.AddText("파 취", xcP3(37.504255, -9.760025), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(40.5, -8.660025)
    Set e = blk.AddText("및", xcP3(39.249388, -14.270877), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(40.447686, -13.170877)
    Set e = blk.AddText("복 구", xcP3(37.504255, -18.560025), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(40.5, -17.460025)
    Set e = blk.AddText("2", xcP3(33.89244, -24.001175), 1)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(34.16478, -23.501175)
    Set e = blk.AddText("2", xcP3(33.89244, -19.601175), 1)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(34.16478, -19.101175)
    Set e = blk.AddText("Con'C", xcP3(48.504255, -12.1), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(51.5, -11)
    Set e = blk.AddText("보도블럭", xcP3(46.706809, -20.9), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(51.5, -19.8)
    Set e = blk.AddText("A.S.P", xcP3(48.504255, -7.7), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(51.5, -6.6)
    Set e = blk.AddText("m", xcP3(32.096308, -34.1), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(32.695457, -33)
    Set e = blk.AddText("m", xcP3(32.096308, -29.7), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(32.695457, -28.6)
    Set e = blk.AddText("m", xcP3(32.096308, -20.9), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(32.695457, -19.8)
    Set e = blk.AddText("m", xcP3(32.096308, -25.3), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(32.695457, -24.2)
    Set e = blk.AddText("m", xcP3(32.096308, -12.1), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(32.695457, -11)
    Set e = blk.AddText("m", xcP3(32.096308, -7.7), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(32.695457, -6.6)
    Set e = blk.AddText("측 점 : No.", xcP3(3.082443, -3.3), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(11.320741, -2.2)
    Set e = blk.AddText("관상단", xcP3(12.905106, -34.1), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(16.5, -33)
    Set e = blk.AddText("관주위", xcP3(12.905106, -29.7), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(16.5, -28.6)
    Set e = blk.AddText("용 수", xcP3(13.504255, -25.3), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(16.5, -24.2)
    Set e = blk.AddText("관   저   고", xcP3(3.810213, -12.1), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(11, -11)
    Set e = blk.AddText("지   반   고", xcP3(3.810213, -7.7), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(11, -6.6)
    Set e = blk.AddText("육 상", xcP3(13.504255, -20.9), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(16.5, -19.8)
    Set e = blk.AddText("되메우기", xcP3(0.706809, -31.9), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(5.5, -30.8)
    Set e = blk.AddText("터 파 기", xcP3(0.706809, -23.1), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(5.5, -22)
    Set e = blk.AddText("m", xcP3(67.062365, -16.5), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(67.661514, -15.4)
    Set e = blk.AddText("A.S.C", xcP3(48.504255, -16.5), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(51.5, -15.4)
    Set e = blk.AddText("2", xcP3(68.89244, -32.801175), 1)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(69.16478, -32.301175)
    Set e = blk.AddText("m", xcP3(67.096308, -34.1), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(67.695457, -33)
    Set e = blk.AddText("동 상 방 지 층", xcP3(37.611915, -34.1), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(46, -33)
    Set e = blk.AddText("m", xcP3(32.096308, -16.5), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(32.695457, -15.4)
    Set e = blk.AddText("토   피   고", xcP3(3.810213, -16.5), 2.2)
    e.Alignment = 10: e.TextAlignmentPoint = xcP3(11, -15.4)

    ' --- 속성 정의 15개 ---
    Set at = blk.AddAttribute(2.2, 0, "측점", xcP3(22.356755, -2.2), "측점", "")
    at.Alignment = 9: at.TextAlignmentPoint = xcP3(22.356755, -2.2)
    Set at = blk.AddAttribute(2.2, 0, "지반고", xcP3(31, -6.6), "지반고", "")
    at.Alignment = 11: at.TextAlignmentPoint = xcP3(31, -6.6)
    Set at = blk.AddAttribute(2.2, 0, "관저고", xcP3(31, -11), "관저고", "")
    at.Alignment = 11: at.TextAlignmentPoint = xcP3(31, -11)
    Set at = blk.AddAttribute(2.2, 0, "토피", xcP3(31, -15.4), "토피", "")
    at.Alignment = 11: at.TextAlignmentPoint = xcP3(31, -15.4)
    Set at = blk.AddAttribute(2.2, 0, "터파기(육상)합계", xcP3(31, -19.8), "터파기(육상)합계", "")
    at.Alignment = 11: at.TextAlignmentPoint = xcP3(31, -19.8)
    Set at = blk.AddAttribute(2.2, 0, "터파기(용수)합계", xcP3(31, -24.2), "터파기(용수)합계", "")
    at.Alignment = 11: at.TextAlignmentPoint = xcP3(31, -24.2)
    Set at = blk.AddAttribute(2.2, 0, "되메우기(관주위)", xcP3(31, -28.6), "되메우기(관주위)", "")
    at.Alignment = 11: at.TextAlignmentPoint = xcP3(31, -28.6)
    Set at = blk.AddAttribute(2.2, 0, "되메우기(관상단)", xcP3(31, -33), "되메우기(관상단)", "")
    at.Alignment = 11: at.TextAlignmentPoint = xcP3(31, -33)
    Set at = blk.AddAttribute(2.2, 0, "ASP포장", xcP3(66, -6.6), "ASP포장", "")
    at.Alignment = 11: at.TextAlignmentPoint = xcP3(66, -6.6)
    Set at = blk.AddAttribute(2.2, 0, "CONC포장", xcP3(66, -11), "CONC포장", "")
    at.Alignment = 11: at.TextAlignmentPoint = xcP3(66, -11)
    Set at = blk.AddAttribute(2.2, 0, "ASC포장", xcP3(66, -15.4), "ASC포장", "")
    at.Alignment = 11: at.TextAlignmentPoint = xcP3(66, -15.4)
    Set at = blk.AddAttribute(2.2, 0, "보도블럭포장", xcP3(66, -19.8), "보도블럭포장", "")
    at.Alignment = 11: at.TextAlignmentPoint = xcP3(66, -19.8)
    Set at = blk.AddAttribute(2.2, 0, "되메우기(모래)", xcP3(66, -24.2), "되메우기(모래)", "")
    at.Alignment = 11: at.TextAlignmentPoint = xcP3(66, -24.2)
    Set at = blk.AddAttribute(2.2, 0, "보조기층", xcP3(66, -28.6), "보조기층", "")
    at.Alignment = 11: at.TextAlignmentPoint = xcP3(66, -28.6)
    Set at = blk.AddAttribute(2.2, 0, "동상방지층", xcP3(66, -33), "동상방지층", "")
    at.Alignment = 11: at.TextAlignmentPoint = xcP3(66, -33)
End Sub

' 도면번호(표제란) : 속성 8개 (원본은 도곽 XREF 위에 이 속성만 얹는다)
Public Sub CreateTitleBlock()
    Dim blk As Object, at As Object
    If BlockExists(pBlkTitle) Then Exit Sub
    On Error Resume Next
    Set blk = doc.Blocks.Add(xcP3(0#, 0#), pBlkTitle)
    If Err.Number <> 0 Then Err.Clear: Exit Sub
    On Error GoTo 0
    Set at = blk.AddAttribute(10, 0, "도면타이틀", xcP3(427.500001, 542.868645), "도면타이틀", "")
    at.Alignment = 10: at.TextAlignmentPoint = xcP3(427.500001, 542.868645)
    Set at = blk.AddAttribute(6.5, 0, "축척H", xcP3(492.871031, 527.350497), "축척H", "")
    Set at = blk.AddAttribute(6.5, 0, "축척V", xcP3(492.871031, 517.192082), "축척V", "")
    Set at = blk.AddAttribute(5, 0, "도면명", xcP3(626.615558, 37.442484), "도면명", "")
    at.Alignment = 10: at.TextAlignmentPoint = xcP3(626.615558, 37.442484)
    Set at = blk.AddAttribute(2, 0, "축척A1", xcP3(792.517503, 46.982608), "축척A1", "")
    at.Alignment = 10: at.TextAlignmentPoint = xcP3(792.517503, 46.982608)
    Set at = blk.AddAttribute(2, 0, "SNO1", xcP3(785.985244, 37.633705), "SNO1", "")
    at.Alignment = 10: at.TextAlignmentPoint = xcP3(785.985244, 37.633705)
    Set at = blk.AddAttribute(2, 0, "SNO2", xcP3(799.107836, 37.633705), "SNO2", "")
    at.Alignment = 10: at.TextAlignmentPoint = xcP3(799.107836, 37.633705)
    Set at = blk.AddAttribute(3, 0, "도면번호", xcP3(777.854994, 28.706447), "도면번호", "")
    at.Alignment = 10: at.TextAlignmentPoint = xcP3(777.854994, 28.706447)
End Sub

' HDot : 치수 보조점 (폭 1.0 의 짧은 선, 삽입축척 0.02)
Public Sub CreateDotBlock()
    Dim blk As Object, e As Object
    Dim a0(0 To 3) As Double
    If BlockExists(pBlkDot) Then Exit Sub
    If Len(pBlkDot) = 0 Then Exit Sub
    On Error Resume Next
    Set blk = doc.Blocks.Add(xcP3(0#, 0#), pBlkDot)
    If Err.Number <> 0 Then Err.Clear: Exit Sub
    a0(0) = 0.5: a0(1) = 0#: a0(2) = -0.5: a0(3) = 0#
    Set e = blk.AddLightWeightPolyline(a0)
    e.ConstantWidth = 1#
    On Error GoTo 0
End Sub

