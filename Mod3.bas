Attribute VB_Name = "Módulo3"
'mejor usar value2 para valores de combobox. cambiar
Public Const rowSEC = 6
Public Const trowSEC = 3
Public Const colSEC = 2
Public Const naSheetSEC = "Sectoriales"
Public Const adSEC = "Sectoriales!B:D"
Public dbSEC  As New Database
Public Const AnoSEC = "C2"

Public Const rowOCU = 7
Public Const trowOCU = 4
Public Const colOCU = 2
Public Const naSheetOCU = "Ocupaciones"
Public Const adOcu = "Ocupaciones!B:T"
Public dbOCU  As New Database
Public Const adIdsectorialesOCU = "E6"
Public Const adSalarioBasico = "$H$2"
Public Const ad13vo = "$J$2"
Public Const ad14vo = "$M$2"
Public Const adFond = "$O$2"
Public Const adAportePersonal = "$R$2"
Public Const adAportePatronal = "$U$2"
Public Const adSalDiar = "M6"
Public Const adsemref = "D6"
Public Const ad14cal = "Q6"
Public Const adapatr = "S6"
Public Const adfonr = "Q6"

Public Const rowTRA = 7
Public Const trowTRA = 4
Public Const colTRA = 2
Public Const naSheetTRA = "Trabajadores"
Public Const adTRA = "Trabajadores!B:AD"
Public dbTRA  As New Database
Public Const adidtra = "B6"
Public Const adEst = "H6"
Public Const adPro = "I6"
Public Const adGua = "J6"
Public Const adOcs = "K6"
Public Const adFecna = "Q6"
Public Const adSex = "S6"
Public Const adEsc = "T6"
Public Const adNac = "U6"
Public Const adExa = "AB6"
Public Const adFim = "AC6"


Public Const rowOBR = 7
Public Const trowOBR = 4
Public Const colOBR = 2
Public Const naSheetOBR = "Obras"
Public Const adOBR = "Obras!B:M"
Public dbOBR  As New Database
Public Const adidtrabObr = "L6"

Public Const rowROL = 7        'SI TIENE 1
Public Const trowROL = 4
Public Const colROL = 2
Public Const naSheetROL = "Roles"
Public Const adOROL = "Roles!B:CW"
Public dbROL  As New Database
Public Const adobraidROL = "$I$2"
Public Const adidrol = "$G$2"

Public Const adfechaROL = "$N$2"
Public Const adCAL = "Calendario!B:AK"
