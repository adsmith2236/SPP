VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EasyInp"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10230
   Icon            =   "easyinp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optShrink 
      Caption         =   "Generate Shrink input"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   2160
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.OptionButton optAsym 
      Caption         =   "Generate Asym40 input"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   2640
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog cdlBox 
      Left            =   9600
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Height          =   615
      Left            =   8400
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox txtOut 
      Height          =   3975
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3960
      Width           =   7455
   End
   Begin VB.Frame fraDataIn 
      Caption         =   "1. File to generate input from"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9975
      Begin VB.OptionButton optLogFile 
         Caption         =   "Gaussian output file"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton optChkPtFile 
         Caption         =   "Gaussian formatted checkpoint file"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   3135
      End
      Begin VB.CommandButton cmdInputFile 
         Caption         =   "Browse"
         Height          =   615
         Left            =   8520
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtInpFile 
         Height          =   375
         Left            =   3720
         TabIndex        =   1
         Top             =   480
         Width           =   4575
      End
   End
   Begin VB.Frame fraDataOut 
      Caption         =   "3. Generate input"
      Height          =   5655
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   9975
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   4680
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   615
         Left            =   8280
         TabIndex        =   18
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   615
         Left            =   8280
         TabIndex        =   8
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblProgress 
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   4440
         Width           =   4335
      End
      Begin VB.Label lblCredits 
         Caption         =   "Credits: K. B. Borisenko and P. D. McCaffrey"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   5280
         Width           =   4455
      End
      Begin VB.Label lblCopy 
         Caption         =   "(c) The University of Edinburgh 2005, 2006, 2007"
         Height          =   255
         Left            =   6240
         TabIndex        =   22
         Top             =   5280
         Width           =   3615
      End
   End
   Begin VB.Frame fraOptions 
      Caption         =   "2. Select options"
      Height          =   1695
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   9975
      Begin VB.Frame frmCubic 
         Height          =   735
         Left            =   6720
         TabIndex        =   19
         Top             =   240
         Width           =   3015
         Begin VB.CheckBox chkCubic 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblCubic 
            Caption         =   "Save cubic constants file (.ccc)"
            Height          =   255
            Left            =   480
            TabIndex        =   20
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame fraInpType 
         Height          =   1215
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtTolerance 
         Height          =   375
         Left            =   9000
         TabIndex        =   12
         Text            =   "0.25"
         ToolTipText     =   "Increase bond tolerance if longer bonds are not found"
         Top             =   1080
         Width           =   735
      End
      Begin VB.OptionButton optSymm 
         Caption         =   "Use symmetry coordinates"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3720
         TabIndex        =   11
         Top             =   960
         Width           =   2295
      End
      Begin VB.OptionButton optInternal 
         Caption         =   "Use internal coordinates"
         Height          =   375
         Left            =   3720
         TabIndex        =   10
         Top             =   480
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.Frame fraCoordType 
         Height          =   1215
         Left            =   3480
         TabIndex        =   15
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblTolerance 
         Caption         =   "Bond tolerance:"
         Height          =   375
         Left            =   7680
         TabIndex        =   13
         Top             =   1200
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InpFileName, OutFileName As String
Dim maxatom As Integer, nfflines, nffconst, nNumbers, nccConst As Long
Dim X(), Y(), z() As Double
Dim number() As Variant
Dim covrad(104) As Double
Dim amass(104) As Double
Dim atom() As Integer
Dim natoms, ndummyatoms As Integer
Dim asymbol(104) As String * 4
Dim str, concst, commentline As String
Dim maxpar As Integer
Dim bond(), angle(), tors(), oopb(), linbend(), shrtors() As Integer
Dim nbonds, nangles, ndangles, noopbs, nlinbends, nshrtors As Integer
Dim GenerateShrink, GenerateAsym As Boolean
Dim UseSymmCoord, UseIntCoord As Boolean
Dim ChkPtInput, LogInput, freq, cubic, FileRead As Boolean
Dim pi, bhr, gr, T, tolr As Double
Dim forcefield() As String, ffield() As Double
Dim cConst() As Double
Dim iVer, iRel As Integer



Private Sub cmdExit_Click()
Unload frmMain
End
End Sub

Private Sub Form_Load()
InpFileName = Empty
Dim maxatom, maxpar, maxff, maxcc As Long
maxatom = 110
maxpar = 1000
maxff = (3 * maxatom * (3 * maxatom + 1)) / 2
maxcc = (3 * maxatom * (1 + 3 * maxatom) * (2 + 3 * maxatom)) / 6
ReDim X(maxatom), Y(maxatom), z(maxatom)

' Dimension the number() and ffiled() arrays according to the maxatom: maxff = 3*natom*(3*natom+1)/2
' Dimenson cConst according to max atom maxcc =  sum[n=1 to 3*natoms] [(n**2 + n) /2] which is the same as
' (3*natoms*(1 + 3*natoms)*(2 + 3*natoms))/6
ReDim number(maxcc), ffield(maxff), cConst(maxcc) ' corresponds to maxatom = 110
ReDim atom(maxatom)
ReDim bond(maxpar, 2), angle(maxpar, 3), tors(maxpar, 4)
ReDim oopb(100, 4), linbend(100, 4), shrtors(maxpar, 8)

lblProgress.Caption = Empty
lblProgress.Refresh
ProgressBar1.Value = 0

' Atomic symbols
asymbol(1) = "H"
asymbol(2) = "He"
asymbol(3) = "Li"
asymbol(4) = "Be"
asymbol(5) = "B"
asymbol(6) = "C"
asymbol(7) = "N"
asymbol(8) = "O"
asymbol(9) = "F"
asymbol(10) = "Ne"
asymbol(11) = "Na"
asymbol(12) = "Mg"
asymbol(13) = "Al"
asymbol(14) = "Si"
asymbol(15) = "P"
asymbol(16) = "S"
asymbol(17) = "Cl"
asymbol(18) = "Ar"
asymbol(19) = "K"
asymbol(20) = "Ca"
asymbol(21) = "Sc"
asymbol(22) = "Ti"
asymbol(23) = "V"
asymbol(24) = "Cr"
asymbol(25) = "Mn"
asymbol(26) = "Fe"
asymbol(27) = "Co"
asymbol(28) = "Ni"
asymbol(29) = "Cu"
asymbol(30) = "Zn"
asymbol(31) = "Ga"
asymbol(32) = "Ge"
asymbol(33) = "As"
asymbol(34) = "Se"
asymbol(35) = "Br"
asymbol(36) = "Kr"
asymbol(37) = "Rb"
asymbol(38) = "Sr"
asymbol(39) = "Y"
asymbol(40) = "Zr"
asymbol(41) = "Nb"
asymbol(42) = "Mo"
asymbol(43) = "Tc"
asymbol(44) = "Ru"
asymbol(45) = "Rh"
asymbol(46) = "Pd"
asymbol(47) = "Ag"
asymbol(48) = "Cd"
asymbol(49) = "In"
asymbol(50) = "Sn"
asymbol(51) = "Sb"
asymbol(52) = "Te"
asymbol(53) = "I"
asymbol(54) = "Xe"
asymbol(55) = "Cs"
asymbol(56) = "Ba"
asymbol(57) = "La"
asymbol(58) = "Ce"
asymbol(59) = "Pr"
asymbol(60) = "Nd"
asymbol(61) = "Pm"
asymbol(62) = "Sm"
asymbol(63) = "Eu"
asymbol(64) = "Gd"
asymbol(65) = "Tb"
asymbol(66) = "Dy"
asymbol(67) = "Ho"
asymbol(68) = "Er"
asymbol(69) = "Tm"
asymbol(70) = "Yb"
asymbol(71) = "Lu"
asymbol(72) = "Hf"
asymbol(73) = "Ta"
asymbol(74) = "W"
asymbol(75) = "Re"
asymbol(76) = "Os"
asymbol(77) = "Ir"
asymbol(78) = "Pt"
asymbol(79) = "Au"
asymbol(80) = "Hg"
asymbol(81) = "Tl"
asymbol(82) = "Pb"
asymbol(83) = "Bi"
asymbol(84) = "Po"
asymbol(85) = "At"
asymbol(86) = "Rn"
asymbol(87) = "Fr"
asymbol(88) = "Ra"
asymbol(89) = "Ac"
asymbol(90) = "Th"
asymbol(91) = "Pa"
asymbol(92) = "U"
asymbol(93) = "Np"
asymbol(94) = "Pu"
asymbol(95) = "Am"
asymbol(96) = "Cm"
asymbol(97) = "Bk"
asymbol(98) = "Cf"
asymbol(99) = "Es"
asymbol(100) = "Fm"
asymbol(101) = "Md"
asymbol(102) = "No"
asymbol(103) = "Lr"
asymbol(104) = "X"

' Covalent radii
covrad(1) = 0.4
covrad(2) = 0.3
covrad(3) = 1.23
covrad(4) = 0.89
covrad(5) = 0.88
covrad(6) = 0.75
covrad(7) = 0.7
covrad(8) = 0.66
covrad(9) = 0.64
covrad(10) = 0.4
covrad(11) = 1.34
covrad(12) = 1.36
covrad(13) = 1.25
covrad(14) = 1.17
covrad(15) = 1.1
covrad(16) = 1.04
covrad(17) = 0.99
covrad(18) = 1.74
covrad(19) = 2.03
covrad(20) = 1.74
covrad(21) = 1.44
covrad(22) = 1.32
covrad(23) = 1.32
covrad(24) = 1.2
covrad(25) = 1.17
covrad(26) = 1.17
covrad(27) = 1.16
covrad(28) = 1.15
covrad(29) = 1.17
covrad(30) = 1.25
covrad(31) = 1.25
covrad(32) = 1.22
covrad(33) = 1.21
covrad(34) = 1.17
covrad(35) = 1.14
covrad(36) = 1.89
covrad(37) = 2.48
covrad(38) = 1.92
covrad(39) = 1.62
covrad(40) = 1.45
covrad(41) = 1.34
covrad(42) = 1.29
covrad(43) = 1.36
covrad(44) = 1.24
covrad(45) = 1.25
covrad(46) = 1.28
covrad(47) = 1.34
covrad(48) = 1.41
covrad(49) = 1.5
covrad(50) = 1.4
covrad(51) = 1.41
covrad(52) = 1.37
covrad(53) = 1.33
covrad(54) = 2.09
covrad(55) = 2.35
covrad(56) = 1.98
covrad(57) = 1.69
covrad(58) = 1.65
covrad(59) = 1.65
covrad(60) = 1.64
covrad(61) = 1.81
covrad(62) = 1.66
covrad(63) = 1.85
covrad(64) = 1.61
covrad(65) = 1.59
covrad(66) = 1.59
covrad(67) = 1.58
covrad(68) = 1.57
covrad(69) = 1.56
covrad(70) = 1.7
covrad(71) = 1.56
covrad(72) = 1.44
covrad(73) = 1.34
covrad(74) = 1.3
covrad(75) = 1.28
covrad(76) = 1.26
covrad(77) = 1.26
covrad(78) = 1.29
covrad(79) = 1.34
covrad(80) = 1.44
covrad(81) = 1.55
covrad(82) = 1.54
covrad(83) = 1.52
covrad(84) = 1.53
covrad(85) = 0#
covrad(86) = 0#
covrad(87) = 2.7
covrad(88) = 2.23
covrad(89) = 0#
covrad(90) = 1.8
covrad(91) = 1.61
covrad(92) = 1.39
covrad(93) = 1.31
covrad(94) = 1.51
covrad(95) = 0#
covrad(96) = 0#
covrad(97) = 0#
covrad(98) = 0#
covrad(99) = 0#
covrad(100) = 0#
covrad(101) = 0#
covrad(102) = 0#
covrad(103) = 0#
covrad(104) = 0#

' Most abundant or stable isotope atomic mass
amass(1) = 1.00783
amass(2) = 4.0026
amass(3) = 7.016
amass(4) = 9.01218
amass(5) = 11.00931
amass(6) = 12#
amass(7) = 14.00307
amass(8) = 15.99491
amass(9) = 18.9984
amass(10) = 19.99244
amass(11) = 22.9898
amass(12) = 23.98504
amass(13) = 26.98154
amass(14) = 27.97693
amass(15) = 30.97376
amass(16) = 31.97207
amass(17) = 34.96885
amass(18) = 39.9624
amass(19) = 38.96371
amass(20) = 39.96259
amass(21) = 44.95592
amass(22) = 45.948
amass(23) = 50.944
amass(24) = 51.9405
amass(25) = 54.9381
amass(26) = 55.9349
amass(27) = 58.9332
amass(28) = 57.9353
amass(29) = 62.9298
amass(30) = 63.9291
amass(31) = 68.9257
amass(32) = 73.9219
amass(33) = 74.9216
amass(34) = 79.9165
amass(35) = 78.9183
amass(36) = 83.912
amass(37) = 84.9117
amass(38) = 87.9056
amass(39) = 88.9054
amass(40) = 89.9043
amass(41) = 92.906
amass(42) = 97.9055
amass(43) = 97.9072
amass(44) = 101.9037
amass(45) = 102.9048
amass(46) = 105.9032
amass(47) = 106.90509
amass(48) = 113.9036
amass(49) = 114.9041
amass(50) = 117.9018
amass(51) = 120.9038
amass(52) = 129.9067
amass(53) = 126.9004
amass(54) = 131.9042
amass(55) = 132.9051
amass(56) = 137.905
amass(57) = 138.9061
amass(58) = 139.9053
amass(59) = 140.9074
amass(60) = 143.9099
amass(61) = 144.9128
amass(62) = 151.9195
amass(63) = 152.9209
amass(64) = 157.9241
amass(65) = 159.925
amass(66) = 163.9288
amass(67) = 164.9303
amass(68) = 165.9304
amass(69) = 168.9344
amass(70) = 173.939
amass(71) = 174.9409
amass(72) = 179.9468
amass(73) = 180.948
amass(74) = 183.951
amass(75) = 186.956
amass(76) = 189.9586
amass(77) = 192.9633
amass(78) = 194.9648
amass(79) = 196.9666
amass(80) = 201.9706
amass(81) = 204.9745
amass(82) = 207.9766
amass(83) = 208.9804
amass(84) = 208.9825
amass(85) = 210.9875
amass(86) = 222.0175
amass(87) = 223.0198
amass(88) = 226.0254
amass(89) = 227.0278
amass(90) = 232.0382
amass(91) = 231.0359
amass(92) = 238.0508
amass(93) = 237.048
amass(94) = 244.0642
amass(95) = 243.0614
amass(96) = 247.0704
amass(97) = 247.0702
amass(98) = 249.0748
amass(99) = 254.0881
amass(100) = 0#
amass(101) = 0#
amass(102) = 0#
amass(103) = 0#
amass(104) = 0#

GenerateAsym = False
GenerateShrink = False
UseSymmCoord = False
UseIntCoord = False

' Physical constants
        pi = 3.1415926536
' Bohr radius
        bhr = 0.529177249
' Room temperature T K
        T = 298.15
' Conversion from radians to degrees
        gr = 180# / pi
        
        For i = 1 To maxpar
            For j = 1 To 2
                bond(i, j) = 0
                angle(i, j) = 0
                tors(i, j) = 0
                shrtors(i, j) = 0
            Next j
            angle(i, 3) = 0
            tors(i, 3) = 0
            tors(i, 4) = 0
        Next i
'Set version and relase
iVer = 2
iRel = 0

frmMain.Caption = "EasyInp v" & CStr(iVer) & "." & CStr(iRel)
Call optInternal_Click
Call optShrink_Click
Call optLogFile_Click



End Sub

Private Sub cmdInputFile_Click()
cdlBox.DialogTitle = "Open file"
cdlBox.InitDir = App.Path
cdlBox.Flags = cdlOFNHideReadOnly
If ChkPtInput Then
cdlBox.Filter = "Formatted checkpoint file (*.fchk)|*.fchk|All (*.*)|*.*"
End If
If LogInput Then
cdlBox.Filter = "Gaussian output file (*.log, *.out)|*.log;*.out|All (*.*)|*.*"
End If
cdlBox.ShowOpen
InpFileName = cdlBox.FileName
txtInpFile.Text = InpFileName
End Sub

Private Sub cmdSave_Click()
Dim OutFFFile, outstr As String, ncol, nlines As Integer
cdlBox.DialogTitle = "Save file"
cdlBox.CancelError = True
cdlBox.Flags = dlgOFHideReadOnly
cdlBox.Filter = "Shrink/Asym40 input file (*.dat)|*.dat|All (*.*)|*.*"
OutFileName = cdlBox.FileName

'On Error GoTo ErrorHandler

pos = InStr(1, OutFileName, ".")
If pos <> 0 And pos = Len(OutFileName) - 3 Then
    OutFileName = Mid$(OutFileName, 1, Len(OutFileName) - 4) & ".dat"
ElseIf pos <> 0 And pos = Len(OutFileName) - 4 Then
    OutFileName = Mid$(OutFileName, 1, Len(OutFileName) - 5) & ".dat"
Else
    OutFileName = OutFileName & ".dat"
End If

cdlBox.FileName = OutFileName

cdlBox.ShowSave

' Get the edited filename
OutFileName = cdlBox.FileName

If OutFileName <> Empty Then
    SaveFile (OutFileName)
    
    If GenerateShrink Then
        ' Save separate force constants file for Shrink
    
        pos = InStr(1, OutFileName, ".")
        If pos <> 0 And pos = Len(OutFileName) - 3 Then
            OutFFFile = Mid$(OutFileName, 1, Len(OutFileName) - 4) & ".ffc"
        Else
            OutFFFile = OutFileName & ".ffc"
        End If
        FileNum = FreeFile
        Open OutFFFile For Output As #FileNum
        
        If cubic Then
            ' Save separate .ccc file for Shrink
            pos = InStr(1, OutFileName, ".")
            If pos <> 0 And pos = Len(OutFileName) - 3 Then
                OutFFFile = Mid$(OutFileName, 1, Len(OutFileName) - 4) & ".ccc"
            Else
                OutFFFile = OutFileName & ".ccc"
            End If
            FileNum2 = FreeFile
            Open OutFFFile For Output As #FileNum2
        End If
    
        If ChkPtInput Then
            For i = 1 To nfflines
                Print #FileNum, forcefield(i)
        Next i
    End If
    
    If LogInput Then
       
        ' How many lines?
        ncol = 5
        If nffconst \ ncol - nffconst / ncol <> 0 Then
            nlines = nffconst \ ncol + 1
        Else
            nlines = nffconst \ ncol
        End If
    
    lblProgress.Caption = "Writting .ffc file..."
    lblProgress.Refresh
    ProgressBar1.Min = 0
    ProgressBar1.Max = nlines
    ProgressBar1.Value = 0
    
    k = 0
    For i = 1 To nlines
        outstr = Empty
        For j = 1 To 5
            k = k + 1
            If k <= nffconst Then
                outstr = outstr & "  " & Format(ffield(k), "0.00000000")
            End If
        
        Next j
        Print #FileNum, outstr
    ProgressBar1.Value = ProgressBar1.Value + 1
    DoEvents
    Next i
    End If
    
    
'    If cubic Then
'        Dim ncc As Integer
'        k = 1
'        For i = 1 To 3 * natoms ' loop over number of k blocks
'            'ncc = (i * i + i) / 2 'number constants to be written for this k block
'            outstr = "K= " & CStr(i) & " block:"
'            Print #FileNum2, outstr
'            outstr = "   "
'            For j = 1 To i
'                outstr = outstr & "     " & CStr(j) & "     "
'            Next j
'            Print #FileNum2, outstr
'            For j = 1 To i
'                outstr = Empty
'                outstr = CStr(j) & " "
'                For l = 1 To j
'                outstr = outstr & " " & Format(cConst(k), "0.00000000")
'                k = k + 1
'                Next l
'                Print #FileNum2, outstr
'            Next j
'        Next i
'        Close (FileNum2)
'    End If

    If cubic Then
        Dim n As Long
        Dim m As Long
        Dim l As Long
        Dim p As Long
        Dim q As Long
        Dim npossum As Long
        
        lblProgress.Caption = "Writting .ccc file..."
        lblProgress.Refresh
        ProgressBar1.Min = 0
        ProgressBar1.Max = 3 * natoms
        ProgressBar1.Value = 0
        
        For n = 1 To 3 * natoms 'Number of k-blocks
            
            outstr = " K=" & Format$(Format$(n, "##0"), "@@@") & " block:"
            Print #FileNum2, outstr
            m = n \ 5 '= number of sections (integer division)
            If n Mod 5 > 0 Then m = m + 1 'adds extra sections if greater than integer (remiander)
            For l = 1 To m 'Loop over sections
                p = 5 * (l - 1) + 1 'Which section are we in?
                For i = p To n 'Loop over section
                    If i <= p + 4 Then
                        q = i
                    Else
                        q = p + 4
                    End If
                    outstr = "   "
                    For j = p To q 'Loop colum vaulues
                        outstr = outstr & "      " & Format$(Format$(j, "#####0"), "@@@@@@@@") 'Print #FileNum2, CStr(((i - 1) * i) / 2 + j) & " "; 'Format(cConst(k), "0.00000000")
                    Next j
                Next i
                Print #FileNum2, outstr
                For i = p To n 'Loop over section
                    If i <= p + 4 Then
                        q = i
                    Else
                        q = p + 4
                    End If
                    outstr = "       "
                    outstr = Empty
                    outstr = "    " & Format$(Format$(i, "##0"), "@@@") & " "
                    For j = p To q 'Loop printing values
                        outstr = outstr & Format$(Format$(cConst(((i - 1) * i) / 2 + j + npossum), "0.0000000000"), "@@@@@@@@@@@@@") & " "
                    Next j
                    Print #FileNum2, outstr
                Next i
            Next l
            npossum = npossum + (n * n + n) / 2
            'MsgBox (npossum)
        ProgressBar1.Value = ProgressBar1.Value + 1
        DoEvents
        Next n
        Close (FileNum2)
    End If
    
    Close (FileNum)
    
    End If
    
End If

lblProgress.Caption = "Done"
lblProgress.Refresh
Exit Sub

ErrorHandler:

    Call ErrorMessage("The file has not been saved!")
    

End Sub

Private Sub SaveFile(thefile As String)
Dim f As Object, stream As Object
Set f = CreateObject("Scripting.FileSystemObject")
Set stream = f.CreateTextFile(thefile)
stream.Write txtOut.Text
stream.Close
End Sub

Private Sub cmdGenerate_Click()
txtOut.Text = Empty
FileRead = False
lblProgress.Caption = Empty
lblProgress.Refresh
ProgressBar1.Value = 0
    
        If GenerateAsym And UseIntCoord And ChkPtInput Then
            Call InitialiseValues
            Call ReadChk
            If FileRead Then
                Call FindIntCoord
                Call GenerateAsymIntCoord
            End If
        End If
        If GenerateShrink And UseIntCoord And ChkPtInput Then
            Call InitialiseValues
            Call ReadChk
            If FileRead Then
                Call FindIntCoord
                Call GenerateShrinkIntCoord
            End If
        End If
        If GenerateAsym And UseIntCoord And LogInput Then
            Call InitialiseValues
            Call ReadLog
            If FileRead Then
                Call FindIntCoord
                Call GenerateAsymIntCoord
            End If
        End If
        If GenerateShrink And UseIntCoord And LogInput Then
            Call InitialiseValues
            Call ReadLog
            If FileRead Then
                Call FindIntCoord
                Call GenerateShrinkIntCoord
            End If
        End If
        If GenerateAsym And UseSymmCoord And ChkPtInput Then
            'Empty
        End If
        If GenerateShrink And UseSymmCoord And ChkPtInput Then
            'Empty
        End If
        If GenerateAsym And UseSymmCoord And LogInput Then
            'Empty
        End If
        If GenerateShrink And UseSymmCoord And LogInput Then
            'Empty
        End If
        lblProgress.Caption = Empty
        lblProgress.Refresh
        ProgressBar1.Value = 0
   
End Sub

Private Sub ReadChk()
Dim concstr As String

On Error GoTo ErrorHandler

' Read in data from input file
FileNum = FreeFile
If InpFileName = Empty Then
    Call cmdInputFile_Click
End If
If InpFileName <> Empty Then
Open InpFileName For Input As #FileNum

' Read the first line as a comment line
        Input #FileNum, commentline

' Find the number of atoms and atomic numbers
Do While Not EOF(FileNum)
        Input #FileNum, str
        If InStr(1, str, "Atomic numbers") <> 0 Then Exit Do
Loop
' Number of atoms
        pos = InStr(1, str, "N=") + 2
        natoms = Val(Mid(str, pos, Len(str)))
' Find the atomic numbers
' How many strings?
        ncol = 6
        If natoms \ ncol - natoms / ncol <> 0 Then
            nstrings = natoms \ ncol + 1
        Else
            nstrings = natoms \ ncol
        End If
        concstr = Empty
        For i = 1 To nstrings
           Line Input #FileNum, str
           concstr = concstr & str & "   "
        Next i
' Find numbers
        Call FindNumbers(concstr)
        For i = 1 To natoms
            atom(i) = number(i)
        Next i
' Find the Cartesian coordinates
Do While Not EOF(FileNum)
        Input #FileNum, str
        If InStr(1, str, "Current cartesian coordinates") Then Exit Do
Loop
' Cartesian coordinates
' How many lines?
        ncol = 5
        If natoms * 3 \ ncol - natoms * 3 / ncol <> 0 Then
            nstrings = natoms * 3 \ ncol + 1
        Else
            nstrings = natoms * 3 \ ncol
        End If
        concstr = Empty
        For i = 1 To nstrings
           Line Input #FileNum, str
           concstr = concstr & str & "   "
        Next i
' Find values of Cartesian coordinates
        Call FindNumbers(concstr)
        j = 1
        For i = 1 To natoms
            X(i) = Val(number(j)) * bhr
            j = j + 1
            Y(i) = Val(number(j)) * bhr
            j = j + 1
            z(i) = Val(number(j)) * bhr
            j = j + 1
        Next i

freq = False

' Read in the force field
Do While Not EOF(FileNum)
        Input #FileNum, str
        If InStr(1, str, "Cartesian Force Constants") Then
            freq = True
            Exit Do
        End If
Loop
        
If Not freq Then
    Call ErrorMessage("The file does not appear to have frequencies calculation!")
    Exit Sub
End If
        
        pos = InStr(1, str, "N=") + 2
        nff = Val(Mid(str, pos, Len(str)))
' How many lines?
        ncol = 5
        If nff \ ncol - nff / ncol <> 0 Then
            nfflines = nff \ ncol + 1
        Else
            nfflines = nff \ ncol
        End If
        ReDim forcefield(nfflines)
'Read in the force field
        For i = 1 To nfflines
           Line Input #FileNum, forcefield(i)
        Next i
Close (FileNum)
FileRead = True
End If

Exit Sub
ErrorHandler:
    Call ErrorMessage("Unidentified file content. Check formatting of the input file!")

End Sub

Private Sub ReadLog()
Dim archline, cartcstr, ffline, ccline As String

On Error GoTo ErrorHandler

' Read in data from input file
FileNum = FreeFile
If InpFileName = Empty Then
    Call cmdInputFile_Click
End If
If InpFileName <> Empty Then

    Open InpFileName For Input As #FileNum

    freq = False
    cubic = False

    ' Look for "Frequencies" word in the output
    Do While Not EOF(FileNum)
        Line Input #FileNum, str
        If InStr(1, str, "Frequencies") <> 0 Then
            freq = True
            Exit Do
        End If
    Loop

    If Not freq Then
        Call ErrorMessage("The file does not appear to have frequencies calculation!")
        Exit Sub
    End If

    ' Look for "Third derivatives" words in the output
    Dim Response As Integer
    Do While Not EOF(FileNum)
        Line Input #FileNum, str
        If InStr(1, str, "Third derivatives") <> 0 And GenerateShrink Then
            Response = MsgBox("" & vbCrLf & "Would you like to generate a .ccc file?", vbInformation + vbYesNo, "Cubic constants detected")
            If Response = vbYes Then
                cubic = True
                chkCubic.Enabled = True
                chkCubic.Value = 1
            End If
            Exit Do
        End If
    Loop


    If Not cubic Then
            chkCubic.Enabled = False
            chkCubic.Value = 0
    End If
    
    Close #FileNum
    Open InpFileName For Input As #FileNum ' rewind the file (in case there was no cubic... would be at end of the file)
    
    ' Look for "Frequencies" word again in the output incase calculation has two parts
    Do While Not EOF(FileNum)
        Line Input #FileNum, str
        If InStr(1, str, "Frequencies") <> 0 Then
            freq = True
            Exit Do
        End If
    Loop
    
    ' Count lines in archive section - This is much faster than reading strings (next section)
    ' and so allows for progress bar when reading concatinating strings... takes ages nad looks like its crashed
    archline = Empty
    Dim nloglines As Long
    Do While Not EOF(FileNum)
        Line Input #FileNum, str
        If nloglines > 0 Then nloglines = nloglines + 1
        If InStr(1, str, "1\1\") <> 0 Then nloglines = 1
        If InStr(1, str, "\\@") <> 0 Then Exit Do
    Loop
    
    Close #FileNum
    Open InpFileName For Input As #FileNum ' rewind the file (in case there was no cubic... would be at end of the file)
    
    ' Look for "Frequencies" word again in the output incase calculation has two parts
    Do While Not EOF(FileNum)
        Line Input #FileNum, str
        If InStr(1, str, "Frequencies") <> 0 Then
            freq = True
            Exit Do
        End If
    Loop
    
    lblProgress.Caption = "Reading archive line..."
    lblProgress.Refresh
    ProgressBar1.Min = 0
    ProgressBar1.Max = nloglines
    ProgressBar1.Value = 0

    

    Dim ncommentLine As Long
    Do While Not EOF(FileNum)
        Line Input #FileNum, str
        If archline <> Empty Then
            archline = archline & LTrim(str)
            ProgressBar1.Value = ProgressBar1.Value + 1
            ncommentLine = ncommentLine + 1
            DoEvents
        End If
        If InStr(1, str, "1\1\") <> 0 Then archline = archline & LTrim(str)
        If InStr(1, str, "\\@") <> 0 Then Exit Do
        If ncommentLine = 5000 Then
                lblProgress.Caption = "Still reading archive line..."
                lblProgress.Refresh
        End If
        If ncommentLine = 10000 Then
                lblProgress.Caption = "Yep... still reading archive line..."
                lblProgress.Refresh
        End If
        If ncommentLine = 15000 Then
                lblProgress.Caption = "I'm a very slow reader..."
                lblProgress.Refresh
        End If
        If ncommentLine = 20000 Then
                lblProgress.Caption = "I wish I was written in Fortran..."
                lblProgress.Refresh
        End If
        If ncommentLine = 25000 Then
                lblProgress.Caption = "Go have a cuppa..."
                lblProgress.Refresh
        End If
        If ncommentLine = 30000 Then
                lblProgress.Caption = "Mines a tea, milk two sugars..."
                lblProgress.Refresh
        End If
    Loop
    
    ' Find the comment line
    pos1 = InStr(1, archline, "\\")
    pos2 = InStr(pos1 + 2, archline, "\\")
    pos3 = InStr(pos2 + 2, archline, "\\")
    commentline = Mid(archline, pos2 + 2, pos3 - pos2 - 2)

    ' Find the number of atoms
    pos4 = InStr(pos3 + 2, archline, "\\")
    natoms = -1
    i = pos3 + 2
    Do While Not i = pos4
        i = InStr(i + 1, archline, "\")
        natoms = natoms + 1
    Loop

    ' Read in the atomic symbols and Cartesian coordinates
    pos = InStr(pos3 + 2, archline, "\")
    cartcstr = Mid(archline, pos + 1, pos4 - pos)

    'cartcstr = RemoveChars(cartcstr, " ")

    For i = 1 To Len(cartcstr)
        If Mid(cartcstr, i, 1) = "," Or Mid(cartcstr, i, 1) = "\" Then
            Mid(cartcstr, i, 1) = " "
        End If
    Next i
    lblProgress.Caption = "Extracting coordinates..."
    lblProgress.Refresh
    ProgressBar1.Min = 0
    ProgressBar1.Max = natoms * 4
    ProgressBar1.Value = 0
    Call FindNumbers(cartcstr)
    j = 1
    For i = 1 To natoms
        For k = 1 To 104
            If RTrim(asymbol(k)) = LTrim(number(j)) Then
                atom(i) = k
                Exit For
            End If
        Next k
        j = j + 1
        X(i) = Val(number(j))
        j = j + 1
        Y(i) = Val(number(j))
        j = j + 1
        z(i) = Val(number(j))
        j = j + 1
    Next i

    ' Read in the force field
    pos5 = InStr(pos4 + 2, archline, "\\")
    pos6 = InStr(pos5 + 2, archline, "\\")
    ffline = Mid(archline, pos5 + 2, pos6 - pos5 - 1)

    'ffline = RemoveChars(ffline, " ")

    For i = 1 To Len(ffline)
        If Mid(ffline, i, 1) = "," Or Mid(ffline, i, 1) = "\" Then
            Mid(ffline, i, 1) = " "
        End If
    Next i
    
    lblProgress.Caption = "Extracting force constants..."
    lblProgress.Refresh
    ProgressBar1.Min = 0
    ProgressBar1.Max = (natoms * 3 * (natoms * 3 + 1)) / 2
    ProgressBar1.Value = 0
    Call FindNumbers(ffline)
    nffconst = nNumbers
    For i = 1 To nffconst
        ffield(i) = Val(number(i))
    Next i
    

    If cubic Then
        ' Read in the force field
        pos7 = InStr(pos6 + 2, archline, "\\")
        pos8 = InStr(pos7 + 2, archline, "\\")
        ccline = Mid(archline, pos7 + 2, pos8 - pos7 - 1)

        'ccline = RemoveChars(ccline, " ")

        For i = 1 To Len(ccline)
            If Mid(ccline, i, 1) = "," Or Mid(ccline, i, 1) = "\" Then
                Mid(ccline, i, 1) = " "
            End If
        Next i

        lblProgress.Caption = "Extracting cubic constants..."
        lblProgress.Refresh
        ProgressBar1.Min = 0
        ProgressBar1.Max = (3 * natoms * (1 + 3 * natoms) * (2 + 3 * natoms)) / 6
        ProgressBar1.Value = 0
        Call FindNumbers(ccline)
        nccConst = nNumbers
        For i = 1 To nccConst
            cConst(i) = Val(number(i))
        Next i
            
    End If

    Close (FileNum)
    FileRead = True
    lblProgress.Caption = "Done"
    lblProgress.Refresh
End If


Exit Sub
ErrorHandler:
    Call ErrorMessage("Error in ReadLog routine..." & vbCrLf & _
    "Make sure the file is in windows and not UNIX form")

End Sub

Private Sub FindIntCoord()
Dim temp1 As Integer

' Read in the bond tolerance
' Increase tolerance tolr if some bonds are not found
tolr = Val(txtTolerance.Text)

' Find bonds
        nbonds = 0
        For i = 1 To natoms - 1
        For j = i + 1 To natoms
        If dist(i, j) <= covrad(atom(i)) + covrad(atom(j)) + tolr Then
            nbonds = nbonds + 1
            bond(nbonds, 1) = i
            bond(nbonds, 2) = j
        End If
        Next j
        Next i

' Find bond angles
        nangles = 0
        For i = 1 To nbonds - 1
        For j = i + 1 To nbonds
        If bond(i, 1) = bond(j, 2) And bond(i, 2) <> bond(j, 1) Then
            nangles = nangles + 1
            angle(nangles, 1) = bond(j, 1)
            angle(nangles, 2) = bond(i, 1)
            angle(nangles, 3) = bond(i, 2)
        End If

        If bond(i, 1) = bond(j, 1) And bond(i, 2) <> bond(j, 2) Then
            nangles = nangles + 1
            angle(nangles, 1) = bond(i, 2)
            angle(nangles, 2) = bond(i, 1)
            angle(nangles, 3) = bond(j, 2)
        End If

        If bond(i, 2) = bond(j, 1) And bond(i, 1) <> bond(j, 2) Then
            nangles = nangles + 1
            angle(nangles, 1) = bond(i, 1)
            angle(nangles, 2) = bond(i, 2)
            angle(nangles, 3) = bond(j, 2)
        End If

        If bond(i, 2) = bond(j, 2) And bond(i, 1) <> bond(j, 1) Then
            nangles = nangles + 1
            angle(nangles, 1) = bond(i, 1)
            angle(nangles, 2) = bond(i, 2)
            angle(nangles, 3) = bond(j, 1)
        End If

        Next j
        Next i

' Find dihedral angles
        ndangles = 0
        For i = 1 To nangles - 1
        For j = i + 1 To nangles

        If angle(i, 2) = angle(j, 1) And angle(i, 1) = angle(j, 2) _
        And angle(i, 3) <> angle(j, 3) Then
            ndangles = ndangles + 1
            tors(ndangles, 1) = angle(i, 3)
            tors(ndangles, 2) = angle(i, 2)
            tors(ndangles, 3) = angle(i, 1)
            tors(ndangles, 4) = angle(j, 3)
        End If

        If angle(i, 2) = angle(j, 1) And angle(i, 3) = angle(j, 2) _
        And angle(i, 1) <> angle(j, 3) Then
            ndangles = ndangles + 1
            tors(ndangles, 1) = angle(i, 1)
            tors(ndangles, 2) = angle(i, 2)
            tors(ndangles, 3) = angle(i, 3)
            tors(ndangles, 4) = angle(j, 3)
        End If

        If angle(i, 1) = angle(j, 2) And angle(i, 2) = angle(j, 3) _
        And angle(i, 3) <> angle(j, 1) Then
            ndangles = ndangles + 1
            tors(ndangles, 1) = angle(j, 1)
            tors(ndangles, 2) = angle(j, 2)
            tors(ndangles, 3) = angle(j, 3)
            tors(ndangles, 4) = angle(i, 3)
        End If
        
        If angle(i, 3) = angle(j, 2) And angle(i, 2) = angle(j, 3) _
        And angle(i, 1) <> angle(j, 1) Then
            ndangles = ndangles + 1
            tors(ndangles, 1) = angle(j, 1)
            tors(ndangles, 2) = angle(j, 2)
            tors(ndangles, 3) = angle(j, 3)
            tors(ndangles, 4) = angle(i, 1)
        End If

        Next j
        Next i

' Harmonize torsions of type x-1-2-x and x-2-1-x in cyclic fragments
        For i = 1 To ndangles - 1
        For j = i + 1 To ndangles
            If tors(i, 2) = tors(j, 3) And tors(i, 3) = tors(j, 2) Then
                temp1 = tors(j, 1)
                tors(j, 1) = tors(j, 4)
                tors(j, 2) = tors(i, 2)
                tors(j, 3) = tors(i, 3)
                tors(j, 4) = temp1
            End If
        Next j
        Next i

End Sub

Private Sub GenerateShrinkIntCoord()
Dim canprint, linear As Boolean
Dim usedtors, neccrd, ntempang As Integer
Dim canusetors() As Boolean
Dim tempang()

' Look for linear fragments of atoms - linear bends
nlinbends = 0
ntempang = 0
ReDim tempang(nangles, 3)
For i = 1 To nangles
    If Abs(180# - ang(angle(i, 1), angle(i, 2), angle(i, 3))) < 5# Then
        nlinbends = nlinbends + 1
        linbend(nlinbends, 1) = angle(i, 1)
        linbend(nlinbends, 2) = angle(i, 2)
        linbend(nlinbends, 3) = angle(i, 3)
        linbend(nlinbends, 4) = 0
    Else
        ntempang = ntempang + 1
        tempang(ntempang, 1) = angle(i, 1)
        tempang(ntempang, 2) = angle(i, 2)
        tempang(ntempang, 3) = angle(i, 3)
    End If
Next i

If nangles = nlinbends Or nbonds = 1 Then
    linear = True
Else
    linear = False
End If

' Remove linear bends from the list of bending coordinates

If nlinbends <> 0 Then
nangles = ntempang
For i = 1 To nangles
    angle(i, 1) = tempang(i, 1)
    angle(i, 2) = tempang(i, 2)
    angle(i, 3) = tempang(i, 3)
Next i
End If

' Look for out of plane bends
noopbs = 0
For i = 1 To nangles - 1
        For j = i + 1 To nangles
            If angle(i, 1) = angle(j, 1) And angle(i, 2) = angle(j, 2) Then
                If ang(angle(i, 1), angle(i, 2), angle(i, 3)) <> 0# And _
                    ang(angle(i, 1), angle(i, 2), angle(i, 3)) <> 180# Then
                    If ang(angle(i, 2), angle(i, 3), angle(j, 3)) <> 0# And _
                    ang(angle(i, 2), angle(i, 3), angle(j, 3)) <> 180# Then
                dh = dihedr(angle(i, 1), angle(i, 2), angle(i, 3), angle(j, 3))
                If Abs(dh) <= 5# Or Abs(Abs(dh) - 180#) <= 5# Then
                    noopbs = noopbs + 1
                    oopb(noopbs, 1) = angle(i, 2)
                    oopb(noopbs, 2) = angle(i, 1)
                    oopb(noopbs, 3) = angle(i, 3)
                    oopb(noopbs, 4) = angle(j, 3)
                End If
                End If
                End If
            End If
        Next j
Next i

' Replace torsions through linear bends with new torsion coordinates

'For i = 1 To ndangles
'Output (tors(i, 1) & "  " & tors(i, 2) & "  " & tors(i, 3) & "  " & tors(i, 4))
'Next i

If ndangles <> 0 And nlinbends <> 0 Then

' Find torsions containing linear bends and non-linear bends
ReDim tempang(ndangles * nlinbends, 4)
ntempang = 0
For i = 1 To ndangles
canprint = False
    For j = 1 To nlinbends
        If tors(i, 1) = linbend(j, 1) And tors(i, 2) = linbend(j, 2) And tors(i, 3) = linbend(j, 3) Or _
           tors(i, 1) = linbend(j, 3) And tors(i, 2) = linbend(j, 2) And tors(i, 3) = linbend(j, 1) Then
            canprint = True
            For k = 1 To nlinbends
                If tors(i, 2) = linbend(k, 1) And tors(i, 3) = linbend(k, 2) And tors(i, 4) = linbend(k, 3) Or _
                   tors(i, 2) = linbend(k, 3) And tors(i, 3) = linbend(k, 2) And tors(i, 4) = linbend(k, 1) Then
                    canprint = False
                End If
            Next k
        End If
    Next j
    If canprint Then
        ntempang = ntempang + 1
        tempang(ntempang, 1) = tors(i, 1)
        tempang(ntempang, 2) = tors(i, 2)
        tempang(ntempang, 3) = tors(i, 3)
        tempang(ntempang, 4) = tors(i, 4)
        canprint = False
    End If
    
    For j = 1 To nlinbends
        If tors(i, 2) = linbend(j, 1) And tors(i, 3) = linbend(j, 2) And tors(i, 4) = linbend(j, 3) Or _
           tors(i, 2) = linbend(j, 3) And tors(i, 3) = linbend(j, 2) And tors(i, 4) = linbend(j, 1) Then
            canprint = True
            For k = 1 To nlinbends
                If tors(i, 1) = linbend(k, 1) And tors(i, 2) = linbend(k, 2) And tors(i, 3) = linbend(k, 3) Or _
                   tors(i, 1) = linbend(k, 3) And tors(i, 2) = linbend(k, 2) And tors(i, 3) = linbend(k, 1) Then
                    canprint = False
                End If
            Next k
        End If
    Next j
    If canprint Then
        ntempang = ntempang + 1
        tempang(ntempang, 1) = tors(i, 1)
        tempang(ntempang, 2) = tors(i, 2)
        tempang(ntempang, 3) = tors(i, 3)
        tempang(ntempang, 4) = tors(i, 4)
        canprint = False
    End If
    
Next i

' Combine torsion containing linear bends
For i = 1 To ntempang - 1
    For j = i + 1 To ntempang
' Linear parts of two torsions intersecting with two common atoms
        If tempang(i, 3) = tempang(j, 1) And tempang(i, 4) = tempang(j, 2) Then
            ndangles = ndangles + 1
            tors(ndangles, 1) = tempang(i, 1)
            tors(ndangles, 2) = tempang(i, 2)
            tors(ndangles, 3) = tempang(j, 3)
            tors(ndangles, 4) = tempang(j, 4)
        End If
        
        If tempang(i, 3) = tempang(j, 4) And tempang(i, 4) = tempang(j, 3) Then
            ndangles = ndangles + 1
            tors(ndangles, 1) = tempang(i, 1)
            tors(ndangles, 2) = tempang(i, 2)
            tors(ndangles, 3) = tempang(j, 2)
            tors(ndangles, 4) = tempang(j, 1)
        End If
        
        If tempang(i, 1) = tempang(j, 3) And tempang(i, 2) = tempang(j, 4) Then
            ndangle = ndangle + 1
            tors(ndangles, 1) = tempang(i, 4)
            tors(ndangles, 2) = tempang(i, 3)
            tors(ndangles, 3) = tempang(j, 2)
            tors(ndangles, 4) = tempang(j, 1)
        End If
        
        If tempang(i, 1) = tempang(j, 2) And tempang(i, 2) = tempang(j, 1) Then
            ndangles = ndangles + 1
            tors(ndangles, 1) = tempang(i, 4)
            tors(ndangles, 2) = tempang(i, 3)
            tors(ndangles, 3) = tempang(j, 3)
            tors(ndangles, 4) = tempang(j, 4)
        End If

' Linear parts of two torsions intersecting with one common atom
'        If tempang(i, 1) = tempang(j, 1) Then
'            ndangles = ndangles + 1
'            tors(ndangles, 1) = tempang(i, 4)
'            tors(ndangles, 2) = tempang(i, 3)
'            tors(ndangles, 3) = tempang(j, 3)
'            tors(ndangles, 4) = tempang(j, 4)
'
'        Output (tors(ndangles, 1) & " " & tors(ndangles, 2) & " " _
'              & tors(ndangles, 3) & " " & tors(ndangles, 4))
'        End If
'        If tempang(i, 1) = tempang(j, 4) Then
'            ndangles = ndangles + 1
'            tors(ndangles, 1) = tempang(i, 4)
'            tors(ndangles, 2) = tempang(i, 3)
'            tors(ndangles, 3) = tempang(j, 2)
'            tors(ndangles, 4) = tempang(j, 1)
'
'        Output (tors(ndangles, 1) & " " & tors(ndangles, 2) & " " _
'              & tors(ndangles, 3) & " " & tors(ndangles, 4))
'        End If
'        If tempang(i, 4) = tempang(j, 1) Then
'            ndangles = ndangles + 1
'            tors(ndangles, 1) = tempang(i, 1)
'            tors(ndangles, 2) = tempang(i, 2)
'            tors(ndangles, 3) = tempang(j, 3)
'            tors(ndangles, 4) = tempang(j, 4)
'
'        Output (tors(ndangles, 1) & " " & tors(ndangles, 2) & " " _
'              & tors(ndangles, 3) & " " & tors(ndangles, 4))
'        End If
'        If tempang(i, 4) = tempang(j, 4) Then
'            ndangles = ndangles + 1
'            tors(ndangles, 1) = tempang(i, 1)
'            tors(ndangles, 2) = tempang(i, 2)
'            tors(ndangles, 3) = tempang(j, 2)
'            tors(ndangles, 4) = tempang(j, 1)
'
'        Output (tors(ndangles, 1) & " " & tors(ndangles, 2) & " " _
'              & tors(ndangles, 3) & " " & tors(ndangles, 4))
'        End If

    Next j
Next i

End If

' Remove torsions including linear bends from the list of torsion coordinates
If ndangles <> 0 And nlinbends <> 0 Then

ReDim tempang(ndangles, 4)
ntempang = 0

For i = 1 To ndangles
    canprint = True
    For j = 1 To nlinbends
        If tors(i, 1) = linbend(j, 1) And tors(i, 2) = linbend(j, 2) And tors(i, 3) = linbend(j, 3) Or _
           tors(i, 2) = linbend(j, 1) And tors(i, 3) = linbend(j, 2) And tors(i, 4) = linbend(j, 3) Or _
           tors(i, 1) = linbend(j, 3) And tors(i, 2) = linbend(j, 2) And tors(i, 3) = linbend(j, 1) Or _
           tors(i, 2) = linbend(j, 3) And tors(i, 3) = linbend(j, 2) And tors(i, 4) = linbend(j, 1) Then
            canprint = False
        End If
    Next j
    If canprint Then
        ntempang = ntempang + 1
        tempang(ntempang, 1) = tors(i, 1)
        tempang(ntempang, 2) = tors(i, 2)
        tempang(ntempang, 3) = tors(i, 3)
        tempang(ntempang, 4) = tors(i, 4)
    End If
Next i

ndangles = ntempang
For i = 1 To ndangles
    tors(i, 1) = tempang(i, 1)
    tors(i, 2) = tempang(i, 2)
    tors(i, 3) = tempang(i, 3)
    tors(i, 4) = tempang(i, 4)
Next i

End If

' Duplicate linear bends and add dummy atoms

ndummyatoms = 0
If nlinbends <> 0 Then
ntempang = nlinbends
' Position of the reference point for dummy atoms
X(maxatom) = 10#
Y(maxatom) = 10#
z(maxatom) = 10#
For i = 1 To nlinbends
    ndummyatoms = ndummyatoms + 1
    Call coor(linbend(i, 2), linbend(i, 1), maxatom, natoms + ndummyatoms, 1#, 0.5 * pi, 0#)
    atom(natoms + ndummyatoms) = 104
    linbend(i, 4) = natoms + ndummyatoms
    
    ndummyatoms = ndummyatoms + 1
    ntempang = ntempang + 1
    Call coor(linbend(i, 2), linbend(i, 1), maxatom, natoms + ndummyatoms, 1#, 0.5 * pi, 0.5 * pi)
    atom(natoms + ndummyatoms) = 104
    linbend(ntempang, 1) = linbend(i, 1)
    linbend(ntempang, 2) = linbend(i, 2)
    linbend(ntempang, 3) = linbend(i, 3)
    linbend(ntempang, 4) = natoms + ndummyatoms
Next i
nlinbends = ntempang
End If

' Convert representation of dihedral angles

For i = 1 To 8
    For j = 1 To maxpar
        shrtors(j, i) = 0
    Next j
Next i

If ndangles <> 0 Then

ReDim canusetors(ndangles)

For i = 1 To ndangles
    canusetors(i) = True
Next i

usedtors = ndangles

nshrtors = 1
shrtors(1, 1) = tors(1, 1)
shrtors(1, 4) = tors(1, 2)
shrtors(1, 5) = tors(1, 3)
shrtors(1, 6) = tors(1, 4)
canusetors(1) = False
usedtors = usedtors - 1

Do While usedtors <> 0

For i = 1 To ndangles
    If canusetors(i) Then
        If shrtors(nshrtors, 4) = tors(i, 2) And shrtors(nshrtors, 5) = tors(i, 3) Then
            canusetors(i) = False
            usedtors = usedtors - 1
            If shrtors(nshrtors, 2) = 0 And shrtors(nshrtors, 1) <> tors(i, 1) Then
                shrtors(nshrtors, 2) = tors(i, 1)
            ElseIf shrtors(nshrtors, 3) = 0 And shrtors(nshrtors, 2) <> tors(i, 1) _
                   And shrtors(nshrtors, 1) <> tors(i, 1) Then
                shrtors(nshrtors, 3) = tors(i, 1)
            End If
            If shrtors(nshrtors, 7) = 0 And shrtors(nshrtors, 6) <> tors(i, 4) Then
                shrtors(nshrtors, 7) = tors(i, 4)
            ElseIf shrtors(nshrtors, 8) = 0 And shrtors(nshrtors, 7) <> tors(i, 4) _
                   And shrtors(nshrtors, 6) <> tors(i, 4) Then
                shrtors(nshrtors, 8) = tors(i, 4)
            End If
        End If
    End If
Next i

For i = 1 To ndangles
    If canusetors(i) Then
        canusetors(i) = False
        usedtors = usedtors - 1
        nshrtors = nshrtors + 1
        shrtors(nshrtors, 1) = tors(i, 1)
        shrtors(nshrtors, 4) = tors(i, 2)
        shrtors(nshrtors, 5) = tors(i, 3)
        shrtors(nshrtors, 6) = tors(i, 4)
        Exit For
    End If
Next i

Loop

End If

' Print the options and configuration sections
nintcoord = nbonds + nangles + nshrtors + noopbs + nlinbends

Output (";options")
Output (";calculations using internal coordinates")
Output (";rotational constants are calculated in MHz")
Output ("form")
Output ("-4")
Output ("rothz")
If cubic Then
    Output ("anharm")
End If
Output (";config")
Output ("scale")
Output (nintcoord)
Output (nintcoord & "*1.")
Output ("local")
Output ("freq")
Output ("5.")

If linear Then
    neccrd = 5
Else
    neccrd = 6
End If

'Print data section
Output ("data")
Output (commentline)
Output ("'C1'" & " 0 " & nintcoord & " 0 0")
Output (natoms & " " & nintcoord & " 0 " & ndummyatoms & " " & neccrd)
Output (nbonds & " " & nangles & " " & noopbs & " " & nlinbends & " " & nshrtors)
Output ("0 " & nintcoord & "*1" & " 0")

' Empty line
        Output (vbCrLf)

' Print Cartesian coordinates and atomic masses
        For i = 1 To natoms + ndummyatoms
            Output ("'" & RTrim(asymbol(atom(i))) & i & "'" & _
            "    " & FormatNumber(amass(atom(i)), 6, vbTrue) & _
            "    " & FormatNumber(X(i), 8, vbTrue) & _
            "    " & FormatNumber(Y(i), 8, vbTrue) & _
            "    " & FormatNumber(z(i), 8, vbTrue))
        Next i

' Empty line
        Output (vbCrLf)

' Print internal coordinates
' Print bond stretches
        For i = 1 To nbonds
            Output (bond(i, 1) & "  " & bond(i, 2))
        Next i
      
' Empty line
        Output (vbCrLf)
      
' Print bond angles
        If nangles <> 0 Then
            For i = 1 To nangles
                Output (angle(i, 1) & "  " & angle(i, 2) & "  " & angle(i, 3))
            Next i
' Empty line
            Output (vbCrLf)
        End If

' Print out of plane bends
        If noopbs <> 0 Then
            For i = 1 To noopbs
                Output (oopb(i, 1) & "  " & oopb(i, 2) & "  " & oopb(i, 3) & "  " & oopb(i, 4))
            Next i
' Empty line
            Output (vbCrLf)
        End If

' Print linear bends
        If nlinbends <> 0 Then
            For i = 1 To nlinbends
                Output (linbend(i, 1) & "  " & linbend(i, 2) & "  " & linbend(i, 3) & "  " & linbend(i, 4))
            Next i
' Empty line
            Output (vbCrLf)
        End If

' Print dihedral angles
        
        If nshrtors <> 0 Then
            For i = 1 To nshrtors
                Output (shrtors(i, 1) & "  " & shrtors(i, 2) & "  " & _
                        shrtors(i, 3) & "  " & shrtors(i, 4) & "  " & _
                        shrtors(i, 5) & "  " & shrtors(i, 6) & "  " & _
                        shrtors(i, 7) & "  " & shrtors(i, 8))
            Next i
        End If
        
' Empty line
        Output (vbCrLf)

' Print the standard room temperature
Output (T)

' Empty line
        Output (vbCrLf)

' Total number of distances and number of bonded distances
        Output (natoms * (natoms - 1) / 2 & " " & nbonds)

' Empty line
        Output (vbCrLf)

' Print the pairs of atoms for which the calculation of the amplitudes is required
' Print the bonds first
        For i = 1 To nbonds
            Output (bond(i, 1) & " " & bond(i, 2))
        Next i

        For i = 1 To natoms - 1
            For j = i + 1 To natoms
                canprint = True
                For k = 1 To nbonds
                    If i = bond(k, 1) And j = bond(k, 2) Or _
                    j = bond(k, 1) And i = bond(k, 2) Then
                        canprint = False
                    End If
                Next k
                If canprint Then Output (i & " " & j)
            Next j
        Next i

fraDataOut.Caption = "SHRINK input file"

p = MsgBox("Shrink input file has been generated. The standard room temperature " & _
        "(298.15K), C1 " & vbCr & "symmetry point group and atomic masses for the " & _
        "most abundant isotopes were used. " & vbCr & "Edit these to " & _
        "correspond to the experimental values. Additional editing may " & _
        "be required" & vbCr & "if the molecule is linear or has cyclic or cage " & _
        "structures." & vbCrLf & vbCr & "Use ""Save"" button to save the " & _
        "Shrink input file and finish generation of the force " & vbCr & "constants file.", vbExclamation, _
        "EasyInp")
End Sub

Private Sub GenerateAsymIntCoord()
Dim canprint As Boolean, outstr As String, nlines, ncol As Integer

' Print the input data
Output (commentline)
Output ("1 " & natoms & " 1 0 0 0 2 0 0")
Output ("0 0 0 0 0")
Output ("0 0 0 0 0 0 0 0 1 0 1")
Output (T)
Output (natoms * 3 - 6)
Output (natoms * 3 & " 0")

' Print force field
If ChkPtInput Then
    For i = 1 To nfflines
        Output (forcefield(i))
    Next i
End If
If LogInput Then
' How many lines?
        ncol = 5
        If nffconst \ ncol - nffconst / ncol <> 0 Then
            nlines = nffconst \ ncol + 1
        Else
            nlines = nffconst \ ncol
        End If

    k = 0
    For i = 1 To nlines
        outstr = Empty
        For j = 1 To 5
            k = k + 1
            If k <= nffconst Then
                outstr = outstr & "  " & Format(ffield(k), "0.00000000")
            End If
        Next j
        Output (outstr)
    Next i
End If

' Empty line
        Output (vbCrLf)

Output (natoms * 3 - 6 & "*1.0")

' Empty line
        Output (vbCrLf)

' Print Cartesian coordinates
        For i = 1 To natoms
            Output ("'" & RTrim(asymbol(atom(i))) & i & "'" & _
            "    " & FormatNumber(X(i), 8, vbTrue) & _
            "    " & FormatNumber(Y(i), 8, vbTrue) & _
            "    " & FormatNumber(z(i), 8, vbTrue))
        Next i

' Empty line
        Output (vbCrLf)

Output (nbonds + nangles + ndangles & " " & 3 * natoms - 6 & " " & nbonds & " " & _
        nangles & " 0 0 2 " & ndangles)

' Empty line
        Output (vbCrLf)

' Print internal coordinates
' Print bond stretches
        For i = 1 To nbonds
            Output (bond(i, 1) & " " & bond(i, 2))
        Next i
      
' Empty line
        Output (vbCrLf)
      
' Print bond angles
        If nangles <> 0 Then
            For i = 1 To nangles
                Output (angle(i, 1) & " " & angle(i, 2) & " " & angle(i, 3))
            Next i
        End If

' Empty line
        Output (vbCrLf)

' Print dihedral angles
        If ndangles <> 0 Then
            For i = 1 To ndangles
                Output (tors(i, 1) & " " & tors(i, 2) & " " & _
                        tors(i, 3) & " " & tors(i, 4))
            Next i
        End If
' Empty line
        Output (vbCrLf)

' Print the standard isotopes atomic masses
        For i = 1 To natoms
            Output (FormatNumber(amass(atom(i)), 6, vbTrue))
        Next i

' Empty line
        Output (vbCrLf)

        Output ("0")

fraDataOut.Caption = "ASYM40 input file"

p = MsgBox("Asym40 input file has been generated. The standard room temperature " & _
        "(298.15K), C1 " & vbCr & "symmetry point group and atomic masses for the " & _
        "most abundant isotopes were used. " & vbCr & "Edit " & _
        "internal coordinates so that their total number is equal " & _
        "to the total number of" & vbCr & "vibrations before running Asym40. Additional editing " & _
        "may be required if the molecule" & vbCr & "is linear or has cyclic or cage " & _
        "structures." & vbCrLf & vbCr & "Use ""Save"" button to save the " & _
        "Asym40 input file.", vbExclamation, _
        "EasyInp")

End Sub

Private Function dist(i, j)
        rr = Sqr((X(i) - X(j)) ^ 2 + (Y(i) - Y(j)) ^ 2 + (z(i) - z(j)) ^ 2)
        dist = rr
End Function
        
Private Sub Output(xx)
    txtOut.Text = txtOut.Text & xx & vbCrLf
End Sub

Private Function acos(ByVal xx As Single)

If xx = 1# Then
    acos = 0
    Exit Function
ElseIf xx = -1# Then
    acos = pi
    Exit Function
Else
    acos = Atn(-xx / Sqr(-xx * xx + 1#)) + 2# * Atn(1#)
End If

End Function

Private Function ang(n1, n2, n3)
         
' Calculate bond angle n1-n2-n3 using Cartesian coordinates

      dd1 = Sqr((X(n1) - X(n2)) ^ 2 + (Y(n1) - Y(n2)) ^ 2 + (z(n1) - z(n2)) ^ 2)
      dd2 = Sqr((X(n2) - X(n3)) ^ 2 + (Y(n2) - Y(n3)) ^ 2 + (z(n2) - z(n3)) ^ 2)
      dd3 = Sqr((X(n1) - X(n3)) ^ 2 + (Y(n1) - Y(n3)) ^ 2 + (z(n1) - z(n3)) ^ 2)
      angcos = (dd1 ^ 2 + dd2 ^ 2 - dd3 ^ 2) / (2# * dd1 * dd2)
      ang = acos(angcos) * 180# / pi

End Function

Private Function dihedr(n1, n2, n3, n4)

' Calculate dihedral angle n1-n2-n3-n4 using Cartesian coordinates

' Bond vectors b, plane normals p

      BXC = X(n1) - X(n2)
      BYC = Y(n1) - Y(n2)
      BZC = z(n1) - z(n2)
      BXK = X(n3) - X(n2)
      BYK = Y(n3) - Y(n2)
      BZK = z(n3) - z(n2)
      PXK = BYC * BZK - BZC * BYK
      PYK = BZC * BXK - BXC * BZK
      PZK = BXC * BYK - BYC * BXK
      BXC = X(n2) - X(n3)
      BYC = Y(n2) - Y(n3)
      BZC = z(n2) - z(n3)
      BXL = X(n4) - X(n3)
      BYL = Y(n4) - Y(n3)
      BZL = z(n4) - z(n3)
      PXL = BYC * BZL - BZC * BYL
      PYL = BZC * BXL - BXC * BZL
      PZL = BXC * BYL - BYC * BXL

' Arccos from plane normals

        QUADL = PXL ^ 2 + PYL ^ 2 + PZL ^ 2
        QUADK = PXK ^ 2 + PYK ^ 2 + PZK ^ 2
        PRODKL = PXL * PXK + PYL * PYK + PZL * PZK
        PDIH = PRODKL / Sqr(QUADL * QUADK)
        If PDIH > 1# Then PDIH = 1#
        If PDIH < -1# Then PDIH = -1#

' Sign of the angle

        PRODBK = PXK * BXL + PYK * BYL + PZK * BZL
        QUADB = BXL ^ 2 + BYL ^ 2 + BZL ^ 2
        PSIGN = PRODBK / Sqr(QUADB * QUADK)
        If PSIGN > 1# Then PSIGN = 1#
        If PSIGN < -1# Then PSIGN = -1#
        PANG = acos(PSIGN)
        If PANG < pi * 0.5 Then PSIGNF = -1#
        If PANG >= pi * 0.5 Then PSIGNF = 1#
        dihedr = PSIGNF * acos(PDIH) * 180# / pi

End Function

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label1_Click()

End Sub

Private Sub optChkPtFile_Click()
ChkPtInput = True
LogInput = False
End Sub

Private Sub optLogFile_Click()
ChkPtInput = False
LogInput = True
End Sub

Private Sub optSymm_Click()
UseSymmCoord = True
UseIntCoord = False
End Sub

Private Sub optInternal_Click()
UseIntCoord = True
UseSymmCoord = False
End Sub

Private Sub optAsym_Click()
fraDataOut.Caption = "3. Click ""Generate"" button to create ASYM40 input file"
GenerateAsym = True
GenerateShrink = False
End Sub

Private Sub optShrink_Click()
fraDataOut.Caption = "3. Click ""Generate"" button to create SHRINK input file"
GenerateShrink = True
GenerateAsym = False
End Sub

Private Sub FindNumbers(ByVal str1 As String)
Dim startpos, nextpos As Long
Dim changed As Boolean
' Delimit is number delimiter
Delimit = " "
If Len(str1) = 0 Then Exit Sub
nNumbers = 0
startpos = 1
nextpos = 1
If Mid$(str1, nextpos, 1) <> Delimit Then
    changed = True
Else
    changed = False
End If
Do While True
    nextpos = nextpos + 1
    If Mid$(str1, nextpos, 1) = Delimit And changed = True Then
        TheNumber = Mid(str1, startpos, nextpos - startpos)
        nNumbers = nNumbers + 1
        If LogInput Then
            ProgressBar1.Value = ProgressBar1.Value + 1
            DoEvents
        End If
        number(nNumbers) = TheNumber
        startpos = nextpos
        changed = False
    End If
    If Mid$(str1, nextpos, 1) <> Delimit And changed = False Then
        changed = True
    End If
    If nextpos = Len(str1) Then
        Exit Do
    End If
Loop
End Sub

Private Function RemoveChars(ByVal pMessage As String, _
pRemovable As String) As String

    Dim lMessage As String
    Dim lCurChar As String
    Dim n As Long
    
    'cycle through message string
    For n = 1 To Len(pMessage)
        'get the current character
        lCurChar = Mid(pMessage, n, 1)
        'if the current character is not in the removable list,
          'add it to the local message
        If InStr(pRemovable, lCurChar) = 0 Then lMessage = _
           lMessage & lCurChar
    Next n
    RemoveChars = lMessage
End Function

Private Sub ErrorMessage(ByVal msg As String)

p = MsgBox(msg, vbCritical, "EasyInp")

End Sub
Private Sub InitialiseValues()

nbonds = 0
nangles = 0
ndangles = 0
noopbs = 0
nlinbends = 0
nshrtors = 0

End Sub


Private Sub coor(IA, IB, IC, n, R, ABN, BETA)

'-------------------------------------------------------
' COORDINATES OF AN ATOM N ARE COMPUTED BY SUBROUTINE COOR
' IF THE COORDINATES OF THREE OTHER ATOMS ARE KNOWN

        COSA = -Cos(ABN) * Cos(BETA)
        SINA = Sqr(1# - COSA * COSA)
        SBETA = Sin(BETA) / SINA
        COSB = Sqr(1# - SBETA * SBETA)
        XAB = X(IB) - X(IA)
        YAB = Y(IB) - Y(IA)
        ZAB = z(IB) - z(IA)
        RAB = Sqr(XAB * XAB + YAB * YAB + ZAB * ZAB)
        XL = XAB / RAB
        YL = YAB / RAB
        ZL = ZAB / RAB
        XAC = X(IC) - X(IA)
        YAC = Y(IC) - Y(IA)
        ZAC = z(IC) - z(IA)
        PROJ = (XAB * XAC + YAB * YAC + ZAB * ZAC) / RAB
        XAP = XL * PROJ
        YAP = YL * PROJ
        ZAP = ZL * PROJ
        XM = XAC - XAP
        YM = YAC - YAP
        ZM = ZAC - ZAP
        RPC = Sqr(XM * XM + YM * YM + ZM * ZM)
        XM = XM / RPC
        YM = YM / RPC
        ZM = ZM / RPC
        XN = YL * ZM - ZL * YM
        YN = ZL * XM - XL * ZM
        ZN = XL * YM - YL * XM
        XSEC = R * COSA
        YSEC = R * SINA * COSB
        ZSEC = R * SINA * SBETA
        X(n) = XL * XSEC + XM * YSEC + XN * ZSEC + X(IA)
        Y(n) = YL * XSEC + YM * YSEC + YN * ZSEC + Y(IA)
        z(n) = ZL * XSEC + ZM * YSEC + ZN * ZSEC + z(IA)

End Sub
