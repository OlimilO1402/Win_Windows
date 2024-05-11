Attribute VB_Name = "MIDC"
Option Explicit

'Public Const IDCANCEL               As Long = 2
'Public Const IDCLOSE                As Long = 8
'Public Const IDCONTINUE             As Long = 11
'
'Public Const IDC_OLEUIHELP          As Long = 99
'Public Const IDC_OFFLINE_HAND       As Long = 103
'
'Public Const IDC_CI_GROUP           As Long = 120
'Public Const IDC_CI_CURRENT         As Long = 121
'Public Const IDC_CI_CURRENTICON     As Long = 122
'Public Const IDC_CI_DEFAULT         As Long = 123
'Public Const IDC_CI_DEFAULTICON     As Long = 124
'Public Const IDC_CI_FROMFILE        As Long = 125
'Public Const IDC_CI_FROMFILEEDIT    As Long = 126
'Public Const IDC_CI_ICONLIST        As Long = 127
'Public Const IDC_CI_LABEL           As Long = 128
'Public Const IDC_CI_LABELEDIT       As Long = 129
'Public Const IDC_CI_BROWSE          As Long = 130
'Public Const IDC_CI_ICONDISPLAY     As Long = 131
'
'Public Const IDC_CV_OBJECTTYPE      As Long = 150
''Public Const IDC_CV_                As Long = 151
'Public Const IDC_CV_DISPLAYASICON   As Long = 152
'Public Const IDC_CV_CHANGEICON      As Long = 153
'Public Const IDC_CV_ACTIVATELIST    As Long = 154
'Public Const IDC_CV_CONVERTTO       As Long = 155
'Public Const IDC_CV_ACTIVATEAS      As Long = 156
'Public Const IDC_CV_RESULTTEXT      As Long = 157
'Public Const IDC_CV_CONVERTLIST     As Long = 158
''Public Const IDC_CV_                As Long = 159
'Public Const IDC_CV_ICONDISPLAY     As Long = 165
'
'Public Const IDC_EL_CHANGESOURCE    As Long = 201
'Public Const IDC_EL_AUTOMATIC       As Long = 202
'Public Const IDC_EL_LINKSLISTBOX    As Long = 206
'Public Const IDC_EL_CANCELLINK      As Long = 209
'Public Const IDC_EL_UPDATENOW       As Long = 210
'Public Const IDC_EL_OPENSOURCE      As Long = 211
'Public Const IDC_EL_MANUAL          As Long = 212
'Public Const IDC_EL_LINKSOURCE      As Long = 216
'Public Const IDC_EL_LINKTYPE        As Long = 217
'Public Const IDC_EL_COL1            As Long = 220
'Public Const IDC_EL_COL2            As Long = 221
'Public Const IDC_EL_COL3            As Long = 222
'
'Public Const IDC_PS_PASTE           As Long = 500
'Public Const IDC_PS_DISPLAYLIST     As Long = 505
'Public Const IDC_PS_DISPLAYASICON   As Long = 506
'Public Const IDC_PS_ICONDISPLAY     As Long = 507
'Public Const IDC_PS_CHANGEICON      As Long = 508
'Public Const IDC_PS_PASTELINK       As Long = 501
'Public Const IDC_PS_PASTELINKLIST   As Long = 504
'Public Const IDC_PS_PASTELIST       As Long = 503
'Public Const IDC_PS_RESULTIMAGE     As Long = 509
'Public Const IDC_PS_RESULTTEXT      As Long = 510
'Public Const IDC_PS_SOURCETEXT      As Long = 502
'
'
'Public Const IDC_BZ_RETRY           As Long = 600
'Public Const IDC_BZ_ICON            As Long = 601
'Public Const IDC_BZ_MESSAGE1        As Long = 602
'Public Const IDC_BZ_SWITCHTO        As Long = 604
'
'Public Const IDC_PU_LINKS           As Long = 900
'Public Const IDC_PU_TEXT            As Long = 901
'Public Const IDC_PU_CONVERT         As Long = 902
'Public Const IDC_PU_ICON            As Long = 908
'
'Public Const IDC_VP_PERCENT         As Long = 1000
'Public Const IDC_VP_CHANGEICON      As Long = 1001
'Public Const IDC_VP_EDITABLE        As Long = 1002
'Public Const IDC_VP_ASICON          As Long = 1003
'Public Const IDC_VP_RELATIVE        As Long = 1005
'Public Const IDC_VP_SPIN            As Long = 1006
'Public Const IDC_VP_ICONDISPLAY     As Long = 1021
'
'Public Const IDC_LP_OPENSOURCE      As Long = 1006
'Public Const IDC_LP_UPDATENOW       As Long = 1007
'Public Const IDC_LP_BREAKLINK       As Long = 1008
'
'Public Const IDC_GP_OBJECTNAME      As Long = 1009
'Public Const IDC_GP_OBJECTTYPE      As Long = 1010
'Public Const IDC_GP_OBJECTSIZE      As Long = 1011
'Public Const IDC_GP_CONVERT         As Long = 1013
'Public Const IDC_GP_OBJECTICON      As Long = 1014
'Public Const IDC_GP_OBJECTLOCATION  As Long = 1022
'
'Public Const IDC_LP_LINKSOURCE      As Long = 1012
'Public Const IDC_LP_CHANGESOURCE    As Long = 1015
'Public Const IDC_LP_AUTOMATIC       As Long = 1016
'Public Const IDC_LP_MANUAL          As Long = 1017
'Public Const IDC_LP_DATE            As Long = 1018
'Public Const IDC_LP_TIME            As Long = 1019
'
'Public Const IDC_UL_METER           As Long = 1029
'Public Const IDC_UL_STOP            As Long = 1030
'Public Const IDC_UL_PERCENT         As Long = 1031
'Public Const IDC_UL_PROGRESS        As Long = 1032
'
'Public Const IDC_VP_RESULTIMAGE     As Long = 1033
'Public Const IDC_VP_SCALETXT        As Long = 1034
'
'
'Public Const IDC_IO_CREATENEW       As Long = 2100
'Public Const IDC_IO_CREATEFROMFILE  As Long = 2101
'Public Const IDC_IO_LINKFILE        As Long = 2102
'Public Const IDC_IO_OBJECTTYPELIST  As Long = 2103
'Public Const IDC_IO_DISPLAYASICON   As Long = 2104
'Public Const IDC_IO_CHANGEICON      As Long = 2105
'Public Const IDC_IO_FILE            As Long = 2106
'Public Const IDC_IO_FILEDISPLAY     As Long = 2107
'Public Const IDC_IO_RESULTIMAGE     As Long = 2108
'Public Const IDC_IO_RESULTTEXT      As Long = 2109
'Public Const IDC_IO_ICONDISPLAY     As Long = 2110
'Public Const IDC_IO_OBJECTTYPETEXT  As Long = 2111
'Public Const IDC_IO_FILETEXT        As Long = 2112
'Public Const IDC_IO_FILETYPE        As Long = 2113
'Public Const IDC_IO_INSERTCONTROL   As Long = 2114
'Public Const IDC_IO_ADDCONTROL      As Long = 2115
'Public Const IDC_IO_CONTROLTYPELIST As Long = 2116
'
'Public Const IDC_STATUS_TITLE       As Long = &H1CF0 '7408
'Public Const IDC_STATUS_DATA1       As Long = &H1CF1 '7409
'Public Const IDC_STATUS_DATA2       As Long = &H1CF2 '7410

