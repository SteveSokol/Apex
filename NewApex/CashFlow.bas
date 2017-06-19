Attribute VB_Name = "CashFlow"
Option Explicit


Public Sub cflow5(job, block)

Dim propna As Integer
Dim prod As Integer
Dim propyr As Integer
Dim maxprodyr As Integer
Dim propmax As Integer
Dim i As Integer
Dim ii As Integer
Dim plastna As Double
Dim firstfract As Double
Dim lastfract As Double
Dim proyrna As Double
Dim life As Double
Dim unitsmined As Double

Erase Secondary
ReDim Secondary(51, 40)
Call finalyear(maxprodyr, propmax)
ReDim VenMan(MaxYr)

BadRor = 0

If job < 4 Then
  For i = 1 To 50
    For ii = 1 To 15
      Ore(i, ii) = 0
    Next ii
  Next i
End If

If job = 1 Then
  unitsmined = 0
  
  For prod = 1 To Np(7)
    If InvYr(prod) > propyr Then propyr = InvYr(prod)
    Call loans(prod)
    Call joint(prod)
  Next prod
  
  For propna = 1 To propmax
    Erase Prop
    propyr = 1
    For prod = 1 To Np(4)
      If Primary(prod, 28) = propna Then
        Call production(job, prod, propna)
        If PLast(prod) > propyr Then propyr = PLast(prod)
      End If
    Next prod
   
    Call propers(propna)
  
  Next propna
    
  Call spreadsheet
  
  For i = 1 To MaxYr
    For ii = 3 To 11
      Secondary(i, ii) = -Secondary(i, ii)
    Next ii
    Secondary(i, 13) = Secondary(i, 13) * -1
    Secondary(i, 14) = Secondary(i, 14) * -1
    Secondary(i, 15) = Secondary(i, 15) * -1
    Secondary(i, 17) = Secondary(i, 17) * -1
    Secondary(i, 22) = Secondary(i, 22) * -1
    Secondary(i, 23) = Secondary(i, 23) * -1
    Secondary(i, 24) = Secondary(i, 24) * -1
    Secondary(i, 27) = Secondary(i, 27) * -1
  Next i
  
  Exit Sub

End If
  
'================== Property Only ====================='
    
If job = 2 Then
  propna = block
  Erase Prop
  For prod = 1 To Np(4)
    If Primary(prod, 28) = propna Then Call production(job, prod, propna)
  Next prod
  
  Call propers(propna)
  
  For i = 1 To MaxYr
    Prop(i, 2) = Prop(i, 2) * -1
    Prop(i, 3) = Prop(i, 3) * -1
    Prop(i, 4) = Prop(i, 4) * -1
    Prop(i, 5) = Prop(i, 5) * -1
    Prop(i, 9) = Prop(i, 9) * -1
    Prop(i, 12) = Prop(i, 12) * -1
  Next i
  
  Exit Sub
    
End If

'==================== Ore Only ======================='

If job > 2 Then
  prod = block
  propna = Primary(block, 28)
  
  Call production(job, prod, propna)
  
  Exit Sub

End If

End Sub

Public Sub production(job, prod, propna)

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim orena As Integer
Dim process As Integer
Dim year1 As Currency
Dim year2 As Currency
Dim life As Double
Dim t As Currency
Dim minable As Double
Dim reserves As Double
Dim unitsmined As Double
Dim proyrna As Double
Dim plastna As Double
Dim firstfract As Double
Dim lastfract As Double
Dim yearna As Double
Dim mining As Currency
Dim milling As Currency
Dim trsm As Currency
Dim refining As Double
Dim deplna As Double
Dim dilute As Currency
Dim grade As Double
Dim convna As Currency
Dim cost As Double
Dim esc As Currency
Dim severance As Currency
Dim tricklife As Double

ReDim VenMan(MaxYr) As Integer

On Error Resume Next

'=============== Production into Ore() Array ==============='

orena = Int(Primary(prod, 27))

If job < 4 Then
  For i = 1 To 51
    For j = 1 To 20
      Ore(i, j) = 0
    Next j
  Next i
End If

tricklife = 0
For i = 1 To 25
  If (Primary(i, 36) > 0 And Primary(i, 40) > 0 And Primary(i, 41) > 0) Then tricklife = 1
Next i

life = FakeLife

If (life = 0 And tricklife = 0) Then Exit Sub

reserves = CDbl(Primary(prod, 38)) / 100
If reserves = 1 Then reserves = 0.99
reserves = reserves / (1 - reserves) * (CDbl(Primary(prod, 36)) * CDbl(Primary(prod, 37)) / 100)
reserves = reserves + (CDbl(Primary(prod, 36)) * CDbl(Primary(prod, 37)) / 100)

unitsmined = unitsmined + reserves

life = CDbl(Primary(prod, 40)) * CDbl(Primary(prod, 41))
If life > 0 Then life = reserves / life

If life = 0 Then
  minable = reserves
Else
  minable = reserves / life
End If

proyrna = CDbl(Primary(prod, 34))

plastna = proyrna + life
firstfract = proyrna - Year(ProYr(prod))
lastfract = 1 - (plastna - Year(PLast(prod)))

'====== Revenue, Operating Cost, and Depletion Calculations ====='

For i = ProYr(prod) To PLast(prod)
  yearna = Incr
  If i = ProYr(prod) Then yearna = yearna - firstfract
  If i = PLast(prod) Then yearna = yearna - lastfract
  If yearna > 1 Then yearna = 1
  If life = 0 Then yearna = 1
  Ore(i, 1) = Ore(i, 1) + minable * yearna
  If job = 4 Then
    t = minable * yearna
  Else
    t = 1
  End If
  mining = 0: milling = 0: trsm = 0: refining = 0: deplna = 0
  For j = 1 To 3
    process = Primary(prod, j + 28)
    If process <> 0 Then
      For k = 1 To 5
        dilute = Primary(orena, 38) / 100
        grade = (1 - dilute) * Primary(orena, k + 1) + dilute * Primary(orena, k + 7)
        convna = Cf(k)
        grade = grade * convna
        grade = grade * Primary(process, k + 53) / 100
        grade = grade - Primary(process, k + 66) * convna / Primary(process, 103)
        If grade < 0 Then grade = 0
        grade = grade * Primary(process, k + 59) / 100
        cost = grade * Primary(process, k + 72)
        esc = Escal(1)
        Call escalate(esc, i, EscYr)
        refining = refining + cost * esc
        grade = grade * Primary(process, k + 79)
        esc = Primary(process, k + 85)
        Call escalate(esc, i, EscYr)
        Ore(i, k + 1) = Ore(i, k + 1) + grade * esc * t
        Ore(i, 7) = Ore(i, 7) + grade * esc * t
        Ore(i, 7 + j) = Ore(i, 7 + j) + grade * esc * t
        deplna = deplna + (grade * esc * (Primary(1, k + 14) / 100))
      Next k
      
'------------------------------- Milling Cost Calculations ---------------------------------'
      
      esc = Escal(1)
      Call escalate(esc, i, EscYr)
      milling = milling + Primary(process, 94) * esc
      Primary(process, 104) = (Primary(process, 101) + Primary(process, 102)) / Primary(process, 103)
      trsm = trsm + Primary(process, 104) * esc
      If Primary(process, 92) > 0 Then
        milling = milling + Primary(process, 93) / Primary(process, 92) * esc
      End If
      For k = 2 To 4
        esc = Escal(1)
        Call escalate(esc, i, EscYr)
        milling = milling + Primary(process, 93 + k) * esc
      Next k
    End If
  Next j
  
'------------------------------- Mining Cost Calculations ---------------------------------'
  
  esc = Escal(1)
  Call escalate(esc, i, EscYr)
  mining = mining + Primary(prod, 46) * esc + Primary(prod, 47) * esc * Primary(prod, 48)
  If Primary(prod, 40) > 0 Then
    mining = mining + Primary(prod, 45) / Primary(prod, 40) * esc
  End If
  
  For k = 2 To 4
    esc = Escal(k)
    Call escalate(esc, i, EscYr)
    mining = mining + Primary(prod, 47 + k) * esc
  Next k

'------------------------------- Operating Cost Summary --------------------------------'

  Ore(i, 11) = Ore(i, 11) + mining * t
  Ore(i, 12) = Ore(i, 12) + milling * t
  Ore(i, 13) = Ore(i, 13) + trsm * t
  Ore(i, 14) = Ore(i, 14) + refining * t
  Ore(i, 15) = Ore(i, 15) + (mining + milling + trsm + refining) * t
  
  year1 = Int(Primary(prod, 116) - Sets(12) + 1)
  year2 = Int(Primary(prod, 117) - Sets(12) + 1)
  
  For k = year1 To year2
    VenMan(k) = Primary(prod, 115) * -1
  Next k
  
  If job < 4 Then
    Prop(i, 1) = Prop(i, 1) + Ore(i, 7) * Ore(i, 1)
    Prop(i, 2) = Prop(i, 2) + (Ore(i, 13) + Ore(i, 14)) * Ore(i, 1)
    Prop(i, 3) = Prop(i, 3) + (Ore(i, 11) + Ore(i, 12) + VenMan(i)) * Ore(i, 1)
    Prop(i, 10) = Prop(i, 10) + Primary(propna, 119) * Ore(i, 1)
    severance = (Ore(i, 7) - Ore(i, 13) - Ore(i, 14)) * Ore(i, 1)
    If severance < 0 Then severance = 0
    Prop(i, 5) = Prop(i, 5) + Sets(4) / 100 * severance
    Prop(i, 5) = Prop(i, 5) + Sets(6) * Ore(i, 1)
    Secondary(i, 8) = Secondary(i, 8) + Sets(4) / 100 * severance
    Secondary(i, 8) = Secondary(i, 8) + Sets(6) * Ore(i, 1)
    If Ore(i, 7) > 0 Then
      deplna = deplna - deplna / Ore(i, 7) * (Ore(i, 13) + Ore(i, 14))
    Else
      deplna = 0
    End If
    
    If deplna < 0 Then deplna = 0
    deplna = deplna * (1 - Primary(propna, 118) / 100) * (1 - Primary(propna, 120) / 100) * (1 - Primary(propna, 121) / 100)
    deplna = deplna * Ore(i, 1)
    Secondary(i, 13) = Secondary(i, 13) + deplna
  End If

Next i

Call depreciate(prod)

Erase VenMan
  
End Sub

Public Sub loans(prod)

Dim i As Integer
Dim j As Integer
Dim n As Integer
Dim ilast As Integer
Dim amt As Currency
Dim esc As Currency
Dim intr As Currency
Dim loan As Currency
Dim loani As Currency
Dim kfactor As Double
InvYr(prod) = Primary(prod, 108) - Sets(12) + 1
amt = Primary(prod, 105) + Primary(prod, 106)
n = Int(Primary(prod, 109) / Incr)
intr = Primary(prod, 107) / 100 * Incr

If amt > 0 And n <> 0 And intr <> 0 Then
  loan = amt * (intr * (1 + intr) ^ n) / ((1 + intr) ^ n - 1)
  ilast = Int((InvYr(prod) - 1) + n)
  ReDim comPrice(ilast) As Currency
  If ilast > 50 Then ilast = 50
  If ilast > MaxYr Then MaxYr = ilast
  
  For i = InvYr(prod) To ilast
    esc = Primary(prod, 86)
    Call escalate(esc, i, InvYr(prod))
    If Primary(prod, 80) * esc > Primary(prod, 110) Then
      comPrice(i) = Primary(prod, 110)
    ElseIf Primary(prod, 80) * esc < Primary(prod, 111) Then
      comPrice(i) = Primary(prod, 111)
    Else
      comPrice(i) = Primary(prod, 80) * esc
    End If
    j = i - InvYr(prod) + 1
    kfactor = 1 - (1 + intr) ^ (-(n - j + 1))
    If Primary(prod, 106) = 0 Then comPrice(i) = 1
    loani = loan * comPrice(i) * kfactor
    Secondary(i, 9) = Secondary(i, 9) + loani
    Secondary(i, 24) = Secondary(i, 24) + ((loan * comPrice(i)) - loani)
  Next i
End If

If ilast > 0 Then
  Secondary(InvYr(prod), 24) = Secondary(InvYr(prod), 24) - (amt * comPrice(InvYr(prod)))
End If
  
End Sub

Public Sub joint(prod)

Dim j As Integer
Dim year1 As Integer
Dim year2 As Integer

InvYr(prod) = Primary(prod, 113) - Sets(12) + 1

Secondary(InvYr(prod), 25) = Secondary(InvYr(prod), 25) + Primary(prod, 112)

year1 = CInt((Primary(prod, 116) - Sets(12)) / Incr + 1)
year2 = CInt((Primary(prod, 117) - Sets(12)) / Incr + 1)

For j = year1 To year2
  Secondary(j, 27) = Primary(prod, 114) / 100
Next j

End Sub

Public Sub propers(propna)

'=================== Royalty Payments ================='

Dim j As Integer
Dim year1 As Integer
Dim year2 As Integer
Dim cap As Currency
Dim pay As Currency
Dim sumpay As Currency

year1 = (Primary(propna, 129) - Sets(12)) / Incr + 1
year2 = (Primary(propna, 130) - Sets(12)) / Incr + 1

'------------------------------------ Advance Payments -----------------------------------'

If Primary(propna, 127) = 1 Then
  For j = year1 To year2
    Prop(j, 11) = Primary(propna, 126) * Incr + Primary(propna, 128) * Incr * (j - year1)
  Next j
End If

'----------------------------------- Minimum Payments -----------------------------------'

If Primary(propna, 127) = 2 Then
  For j = year1 To year2
    Prop(j, 16) = Primary(propna, 126) * Incr + Primary(propna, 128) * Incr * (j - year1)
  Next j
End If

'------------------------------------- Lease Payments --------------------------------------'

If Primary(propna, 127) = 3 Then
  For j = year1 To year2
    Prop(j, 15) = Prop(j, 15) + Primary(propna, 126) * Incr + Primary(propna, 128) * Incr * (j - year1)
  Next j
End If

'----------------------------------- Net Profits Interest -------------------------------------'

For j = 1 To MaxYr
  If Primary(propna, 121) = 0 Or Primary(propna, 122) = 2 Or Primary(propna, 122) = 3 Then Prop(j, 4) = 0
  If Primary(propna, 122) = 3 Then Prop(j, 8) = 0
  Prop(j, 6) = Prop(j, 1) - Prop(j, 2) - Prop(j, 3) - Prop(j, 4) - Prop(j, 5)
  Prop(j, 7) = Prop(j, 6) * Primary(propna, 121) / 100
  If Prop(j, 7) < 0 Then Prop(j, 7) = 0
  Secondary(j, 36) = Prop(j, 7)
  Secondary(j, 37) = Prop(j, 8)
Next j

Call recapture

'----------------------------------- Net Smelter Return -------------------------------------'

For j = 1 To MaxYr
  Prop(j, 9) = Secondary(j, 38)
  Prop(j, 10) = Prop(j, 10) + Secondary(j, 39)
  Prop(j, 10) = Prop(j, 10) + (Prop(j, 1) - Prop(j, 2)) * Primary(propna, 118) / 100
  Prop(j, 10) = Prop(j, 10) + (Prop(j, 1) - Prop(j, 2)) * Primary(propna, 120) / 100
  Secondary(j, 36) = Prop(j, 10)
  Secondary(j, 37) = Prop(j, 11)
Next j

Call recapture

For j = 1 To MaxYr
  Prop(j, 12) = Secondary(j, 38)
  Prop(j, 13) = Prop(j, 10) + Prop(j, 11) - Prop(j, 12) + Prop(j, 15)
  If Prop(j, 13) < Prop(j, 16) Then Prop(j, 13) = Prop(j, 16)
Next j

cap = Primary(propna, 125)

If cap > 0 Then
  For j = 1 To MaxYr
    pay = Prop(j, 13)
    If cap - pay >= 0 Then
      pay = pay
    Else
      pay = cap
    End If
    cap = cap - pay
    Prop(j, 13) = pay
  Next j
End If

sumpay = 0

For j = 1 To MaxYr
  sumpay = sumpay + Prop(j, 13)
  Prop(j, 14) = sumpay
Next j

For j = 1 To MaxYr
  Secondary(j, 1) = Secondary(j, 1) + Prop(j, 1)
  Secondary(j, 3) = Secondary(j, 3) + Prop(j, 13)
  Secondary(j, 4) = Secondary(j, 4) + Prop(j, 2) + Prop(j, 3)
  If Primary(propna, 121) = 0 Then
    Prop(j, 3) = 0
    Prop(j, 4) = 0
    Prop(j, 5) = 0
    Prop(j, 6) = 0
  End If
Next j

End Sub

Public Sub recapture()

Dim i As Integer
Dim sumamount As Currency

'================== Recapture Notes ==================='

'             Secondary( ,36) = payments due (remains untouched)
'             Secondary( ,37) => advance payments
'             Secondary( ,38) is display for recapture in year it occurs
'             Secondary( ,39) is net schedule

'             send Secondary( ,36) and Secondary( , 37) and get back Secondary( ,38) and Secondary( ,39)


For i = 1 To MaxYr
  Secondary(i, 38) = 0
  Secondary(i, 39) = Secondary(i, 36)
Next i

sumamount = 0

For i = 1 To MaxYr
  sumamount = sumamount + Secondary(i, 37)
  If sumamount > 0 Then
    If Secondary(i, 39) - sumamount > 0 Then
      Secondary(i, 38) = Secondary(i, 38) + sumamount
      Secondary(i, 39) = Secondary(i, 39) - sumamount
      sumamount = 0
    Else
      Secondary(i, 38) = Secondary(i, 38) + Secondary(i, 39)
      sumamount = sumamount - Secondary(i, 39)
      Secondary(i, 39) = 0
    End If
  End If
Next i

End Sub

Public Sub escalate(esc, i, eyr)

  If EscMode = 1 Then
    esc = (1 + esc / 100) ^ (Year(i) - eyr)
  Else
    esc = 1 + (esc / 100) * (Year(i) - eyr)
  End If

End Sub

Public Sub spreadsheet()

Dim i As Integer
Dim j As Integer
Dim sumjv As Currency
Dim sumeq As Currency
Dim sevr As Currency
Dim sumcost As Currency
Dim taxable As Currency
Dim mintax As Currency
Dim tempbad As Integer

For i = 0 To 50
  Cash(i) = 0
Next i

For i = 1 To MaxYr
  sevr = (Secondary(i, 1) - Secondary(i, 3) - Secondary(i, 4)) * Sets(5) / 100
  If sevr < 0 Then sevr = 0
  Secondary(i, 8) = Secondary(i, 8) + sevr
  sumcost = 0
  For j = 3 To 11
    sumcost = sumcost + Secondary(i, j)
  Next j
  Secondary(i, 12) = (Secondary(i, 1) + Secondary(i, 2)) - sumcost
Next i
  
Call deplete

For i = 1 To MaxYr
  If PType < 1 Then
    If i > 1 Then
      If Secondary(i - 1, 16) < 0 Then
        Secondary(i, 14) = Secondary(i - 1, 16) * -1
        Secondary(i, 21) = Secondary(i, 14)
      End If
    End If
  End If
  taxable = Secondary(i, 12) - Secondary(i, 13) - Secondary(i, 14)
  Secondary(i, 15) = taxable * Sets(3) / 100
  If PType < 1 Then
    If Secondary(i, 15) < 0 Then Secondary(i, 15) = 0
  End If
  If Secondary(i, 15) > 0 Then
    Secondary(i, 16) = taxable - Secondary(i, 15)
  Else
    Secondary(i, 16) = taxable
  End If
  Secondary(i, 17) = Secondary(i, 16) * Sets(2) / 100
  mintax = (Secondary(i, 16) + Secondary(i, 13)) * 0.2
  If Sets(2) < 20 Then mintax = 0
  If Secondary(i, 17) < mintax Then Secondary(i, 17) = mintax
  If PType < 1 Then
    If Secondary(i, 17) < 0 Then Secondary(i, 17) = 0
  End If
  Secondary(i, 18) = Secondary(i, 16) - Secondary(i, 17)
  sumcost = Secondary(i, 19) + Secondary(i, 20) + Secondary(i, 21) - Secondary(i, 22)
  sumcost = sumcost - Secondary(i, 23) - Secondary(i, 24) + Secondary(i, 25)
  Secondary(i, 26) = Secondary(i, 18) + sumcost
  Secondary(i, 27) = Secondary(i, 27) * Secondary(i, 26)
Next i

For i = 1 To MaxYr
  Secondary(i, 28) = Secondary(i, 26) - Secondary(i, 27)
  If Secondary(i, 28) > 0 Then tempbad = 2
  If Secondary(i, 28) < 0 And tempbad = 2 Then
    BadRor = 2
  End If
  Cash(i) = Secondary(i, 28)
  Secondary(i, 29) = Secondary(i - 1, 29) + Secondary(i, 28)
Next i
  
End Sub

Public Sub deplete()

Dim i As Integer
Dim sumvalue As Currency
Dim amt As Currency
Dim cost As Currency

For i = 1 To MaxYr
  sumvalue = sumvalue + Secondary(i, 1)
Next i

For i = 1 To MaxYr
  amt = Secondary(i, 12) * 0.5
  If amt < 0 Then amt = 0
  If Secondary(i, 13) > amt Then Secondary(i, 13) = amt
  If sumvalue > 0 Then
    cost = AcqiCost * (Secondary(i, 1) / sumvalue)
  Else
    cost = 0
  End If
  If cost > Secondary(i, 13) Then Secondary(i, 13) = cost
  If Primary(1, 15) = 0 Then Secondary(i, 13) = 0
  Secondary(i, 20) = Secondary(i, 13)
Next i

If Sets(2) = 0 Then
  For i = 1 To MaxYr
    Secondary(i, 13) = 0
    Secondary(i, 20) = Secondary(i, 13)
  Next i
End If

End Sub

Public Sub depreciate(prod)

Dim f As Integer
Dim n As Integer
Dim X As Integer
Dim z As Integer
Dim when As Integer
Dim argh As Integer
Dim jeez As Integer
Dim maxprodyr As Integer
Dim propmax As Integer
Dim lower As Integer
Dim upper As Integer
Dim rate As Single
Dim switch As Single
Dim adjust As Currency
Dim basis As Currency
Dim salvtemp As Currency
ReDim Acrs(NumCap, 50) As Currency
ReDim Stln(NumCap, 50) As Currency
ReDim Dcbl(NumCap, 50) As Currency
ReDim Ddbl(NumCap, 50) As Currency
ReDim Defx(NumCap, 50) As Currency
ReDim Soyd(NumCap, 50) As Currency
ReDim Ammt(NumCap, 50) As Currency
ReDim Unop(NumCap, 50) As Currency
ReDim Deve(NumCap, 50) As Currency
ReDim Devd(NumCap, 50) As Currency
ReDim Expn(NumCap, 50) As Currency
ReDim Wrcp(NumCap, 50) As Currency
ReDim Salv(NumCap, 50) As Currency
ReDim Recl(NumCap, 50) As Currency
ReDim Totl(NumCap, 50) As Currency
ReDim Deprecy(NumCap, 50) As Currency

Call finalyear(maxprodyr, propmax)

AcqiCost = 0

For n = 0 To NumCap - 1
  If CapitalData(n).DepPeriod > 0 Or CapitalData(n).DepMethod = "wor" Or CapitalData(n).DepMethod = "exp" Or CapitalData(n).DepMethod = "rec" Or CapitalData(n).DepMethod = "acq" Then
    Select Case CapitalData(n).DepMethod
                  
'========= Modified Accelerated Cost Recovery System ========'
            
      Case "mod", "mac"
        basis = CapitalData(n).PurchaseAmount
        adjust = 2
        If CapitalData(n).DepPeriod > 14 Then adjust = 1.5
        rate = adjust / CapitalData(n).DepPeriod
        For argh = CapitalData(n).DepPeriod To 1 Step -1
          switch = 1 / (argh + 0.5)
          If switch >= rate Then
            jeez = CapitalData(n).DepPeriod - (argh + 1)
            argh = 1
          End If
        Next argh
        For X = 0 To CapitalData(n).DepPeriod
          when = CapitalData(n).InvestYear + X
          z = 1
          If (X = 0) Or (X = CapitalData(n).DepPeriod) Then
            z = 2
          End If
          If when < maxprodyr Then
            Acrs(n, when) = ((basis * adjust) / (CapitalData(n).DepPeriod * z))
            salvtemp = 0
            Deprecy(n, when) = Acrs(n, when)
          Else
            Acrs(n, maxprodyr) = Acrs(n, maxprodyr) + ((basis * adjust) / (CapitalData(n).DepPeriod * z))
            salvtemp = ((basis * adjust) / (CapitalData(n).DepPeriod * z))
            Deprecy(n, maxprodyr) = Acrs(n, maxprodyr)
        End If
          Call autosalvage(n, salvtemp, maxprodyr, Salv())
          If (X <= jeez And when <= 50) Then basis = basis - Acrs(n, when)
        Next X
        when = CapitalData(n).InvestYear
        Totl(n, when) = CapitalData(n).PurchaseAmount
        
'================ Straight Line Depreciation ==============='
      
      Case "str"
        For when = CapitalData(n).InvestYear To (CapitalData(n).InvestYear + CapitalData(n).DepPeriod)
          z = 1
          If when = CapitalData(n).InvestYear Or when = (CapitalData(n).InvestYear + CapitalData(n).DepPeriod) Then z = 2
          If when < maxprodyr Then
            Stln(n, when) = CapitalData(n).PurchaseAmount / (CapitalData(n).DepPeriod * z)
            salvtemp = 0
            Deprecy(n, when) = Stln(n, when)
          Else
            Stln(n, maxprodyr) = Stln(n, maxprodyr) + (CapitalData(n).PurchaseAmount / (CapitalData(n).DepPeriod * z))
            salvtemp = (CapitalData(n).PurchaseAmount / (CapitalData(n).DepPeriod * z))
            Deprecy(n, maxprodyr) = Stln(n, maxprodyr)
          End If
          Call autosalvage(n, salvtemp, maxprodyr, Salv())
        Next when
        when = CapitalData(n).InvestYear
        Totl(n, when) = CapitalData(n).PurchaseAmount
                
'================== Diminishing Balance ================='
      
      Case "dim"
        basis = CapitalData(n).PurchaseAmount
        For when = CapitalData(n).InvestYear To (CapitalData(n).InvestYear + (CapitalData(n).DepPeriod - 1))
          z = 1
          If basis = CapitalData(n).PurchaseAmount Then z = 2
          If when < maxprodyr Then
            Dcbl(n, when) = (basis * (CapitalData(n).DmRate / (z * 100)))
            salvtemp = 0
            Deprecy(n, when) = Dcbl(n, when)
          Else
            Dcbl(n, maxprodyr) = Dcbl(n, maxprodyr) + (basis * (CapitalData(n).DmRate / (z * 100)))
            salvtemp = (basis * (CapitalData(n).DmRate / (z * 100)))
            Deprecy(n, maxprodyr) = Dcbl(n, maxprodyr)
        End If
          Call autosalvage(n, salvtemp, maxprodyr, Salv())
          If when <= 50 Then basis = basis - Dcbl(n, when)
        Next when
        when = CapitalData(n).InvestYear
        Totl(n, when) = CapitalData(n).PurchaseAmount
        
'================ Double Declining Balance ==============='
      
      Case "dou"
        basis = CapitalData(n).PurchaseAmount
        For when = CapitalData(n).InvestYear To (CapitalData(n).InvestYear + (CapitalData(n).DepPeriod - 1))
          z = 1
          If basis = CapitalData(n).PurchaseAmount Then z = 2
          If when < maxprodyr Then
            Ddbl(n, when) = (basis * 2) / (CapitalData(n).DepPeriod * z)
            salvtemp = 0
            Deprecy(n, when) = Ddbl(n, when)
          Else
            Ddbl(n, maxprodyr) = Ddbl(n, maxprodyr) + ((basis * 2) / (CapitalData(n).DepPeriod * z))
            salvtemp = ((basis * 2) / (CapitalData(n).DepPeriod * z))
            Deprecy(n, maxprodyr) = Ddbl(n, maxprodyr)
          End If
          Call autosalvage(n, salvtemp, maxprodyr, Salv())
          If when <= 50 Then basis = basis - Ddbl(n, when)
        Next when
        when = CapitalData(n).InvestYear
        Totl(n, when) = CapitalData(n).PurchaseAmount
        
'====== Deferred Modified Accelerated Cost Recovery System ====='
      
      Case "def"
        basis = CapitalData(n).PurchaseAmount
        adjust = 2
        If CapitalData(n).DepPeriod > 14 Then adjust = 1.5
        rate = adjust / CapitalData(n).DepPeriod
        For argh = CapitalData(n).DepPeriod To 1 Step -1
          switch = 1 / (argh + 0.5)
          If switch >= rate Then
            jeez = CapitalData(n).DepPeriod - (argh + 1)
            argh = 1
          End If
        Next argh
        For X = 0 To CapitalData(n).DepPeriod
          when = CapitalData(n).InvestYear + CapitalData(n).DmRate + X
          z = 1
          If X = 0 Or X = CapitalData(n).DepPeriod Then
            z = 2
          End If
          If when < maxprodyr Then
            Defx(n, when) = ((basis * adjust) / (CapitalData(n).DepPeriod * z))
            salvtemp = 0
            Deprecy(n, when) = Defx(n, when)
          Else
            Defx(n, maxprodyr) = Defx(n, maxprodyr) + ((basis * adjust) / (CapitalData(n).DepPeriod * z))
            salvtemp = ((basis * adjust) / (CapitalData(n).DepPeriod * z))
            Deprecy(n, maxprodyr) = Defx(n, maxprodyr)
          End If
          Call autosalvage(n, salvtemp, maxprodyr, Salv())
          If (X <= jeez And when <= 50) Then basis = basis - Defx(n, when)
        Next X
        when = CapitalData(n).InvestYear
        Totl(n, when) = CapitalData(n).PurchaseAmount
        
'================ Sum of the Years Digits ================'
      
      Case "sum"
        lower = Int((CapitalData(n).DepPeriod * (CapitalData(n).DepPeriod + 1)) / 2)
        z = 0
        For when = CapitalData(n).InvestYear To (CapitalData(n).InvestYear + (CapitalData(n).DepPeriod - 1))
          upper = Int(CapitalData(n).DepPeriod - z)
          If lower > 0 Then
            If when < maxprodyr Then
              Soyd(n, when) = CapitalData(n).PurchaseAmount * (upper / lower)
              salvtemp = 0
              Deprecy(n, when) = Soyd(n, when)
            Else
              Soyd(n, maxprodyr) = Soyd(n, maxprodyr) + (CapitalData(n).PurchaseAmount * (upper / lower))
              salvtemp = (CapitalData(n).PurchaseAmount * (upper / lower))
              Deprecy(n, maxprodyr) = Soyd(n, maxprodyr)
            End If
            Call autosalvage(n, salvtemp, maxprodyr, Salv())
          End If
          z = z + 1
        Next when
        when = CapitalData(n).InvestYear
        Totl(n, when) = CapitalData(n).PurchaseAmount
        
'==================== Amortization ===================='
      
      Case "amo"
        For when = CapitalData(n).InvestYear To (CapitalData(n).InvestYear + (CapitalData(n).DepPeriod - 1))
          If when < maxprodyr Then
            Ammt(n, when) = CapitalData(n).PurchaseAmount / CapitalData(n).DepPeriod
            salvtemp = 0
            Deprecy(n, when) = Ammt(n, when)
          Else
            Ammt(n, maxprodyr) = Ammt(n, maxprodyr) + (CapitalData(n).PurchaseAmount / CapitalData(n).DepPeriod)
            salvtemp = (CapitalData(n).PurchaseAmount / CapitalData(n).DepPeriod)
            Deprecy(n, maxprodyr) = Ammt(n, maxprodyr)
          End If
        Next when
        when = CapitalData(n).InvestYear
        Totl(n, when) = CapitalData(n).PurchaseAmount
        
'=================== Units of Production ================='
      
      Case "uni"
        If Primary(prod, 36) > 0 Then
          For when = Primary(prod, 34) To maxprodyr - 1
            rate = (Primary(prod, 40) * Primary(prod, 41)) / Primary(prod, 36)
            Unop(n, when) = CapitalData(n).PurchaseAmount * rate
            Deprecy(n, when) = Unop(n, when)
          Next when
        End If
        when = CapitalData(n).InvestYear
        Totl(n, when) = CapitalData(n).PurchaseAmount
        
'==================== Development ===================='
      
      Case "dev"
        For when = CapitalData(n).InvestYear To (CapitalData(n).InvestYear + (CapitalData(n).DepPeriod - 1))
          If when = CapitalData(n).InvestYear And Sets(2) <> 0 Then
            Deve(n, when) = CapitalData(n).PurchaseAmount * 0.7
            Totl(n, when) = CapitalData(n).PurchaseAmount * 0.3
            Prop(when, 8) = CapitalData(n).PurchaseAmount
          ElseIf when = CapitalData(n).InvestYear And Sets(2) = 0 Then
            Totl(n, when) = CapitalData(n).PurchaseAmount
            Prop(when, 8) = CapitalData(n).PurchaseAmount
          End If
          If when < maxprodyr Then
            Devd(n, when) = (CapitalData(n).PurchaseAmount * 0.3) / CapitalData(n).DepPeriod
            Deprecy(n, when) = Deve(n, when) + Devd(n, when)
          Else
            Devd(n, maxprodyr) = Devd(n, maxprodyr) + ((CapitalData(n).PurchaseAmount * 0.3) / CapitalData(n).DepPeriod)
            Deprecy(n, maxprodyr) = Devd(n, maxprodyr)
         End If
        Next when

'=================== Working Capital ==================='
   
      Case "wor"
        when = CapitalData(n).InvestYear
        Wrcp(n, when) = CapitalData(n).PurchaseAmount
        Wrcp(n, maxprodyr) = CapitalData(n).PurchaseAmount * -1
        Deprecy(n, when) = Wrcp(n, when)
        
'===================== Expensed ====================='

      Case "exp"
        when = CapitalData(n).InvestYear
        Expn(n, when) = CapitalData(n).PurchaseAmount
        Deprecy(n, when) = Expn(n, when)
        
'==================== Reclamation ====================='
      
      Case "rec"
        when = CapitalData(n).InvestYear
        Recl(n, when) = CapitalData(n).PurchaseAmount
        Deprecy(n, when) = Recl(n, when)
        
'==================== Amortization ===================='

      Case "acq"
        when = CapitalData(n).InvestYear
        AcqiCost = AcqiCost + CapitalData(n).PurchaseAmount
        Totl(n, when) = CapitalData(n).PurchaseAmount
        Deprecy(n, when) = Totl(n, when)
    End Select

'====================== Salvage ======================'

    If CapitalData(n).SalvageAmount > 0 Then
      when = CapitalData(n).SoldYear
    Else
      when = maxprodyr
    End If
    If CapitalData(n).Changed = True And CapitalData(n).SalvageAmount > 0 Then
      Salv(n, when) = CapitalData(n).SalvageAmount
    End If
  End If
Next n

'==================== Calculations ===================='

For when = 1 To MaxYr
  Secondary(when, 2) = 0
  Secondary(when, 5) = 0
  Secondary(when, 6) = 0
  Secondary(when, 11) = 0
  Secondary(when, 19) = 0
  Secondary(when, 22) = 0
  Secondary(when, 23) = 0
  Prop(when, 4) = 0
Next when

For when = 1 To MaxYr
  For n = 0 To NumCap - 1
    Secondary(when, 2) = Secondary(when, 2) + Salv(n, when)
    Secondary(when, 5) = Secondary(when, 5) + Deve(n, when) + Expn(n, when)
    If Sets(2) = 0 Then
      Secondary(when, 6) = Secondary(when, 6) + 0
      Prop(when, 4) = Prop(when, 4) + 0
    Else
      Secondary(when, 6) = Secondary(when, 6) + Acrs(n, when) + Stln(n, when) + Dcbl(n, when) + Ddbl(n, when) + Defx(n, when) + Soyd(n, when) + Ammt(n, when) + Unop(n, when) + Devd(n, when)
      Prop(when, 4) = Prop(when, 4) + Acrs(n, when) + Stln(n, when) + Dcbl(n, when) + Ddbl(n, when) + Defx(n, when) + Soyd(n, when) + Ammt(n, when) + Unop(n, when) + Devd(n, when)
    End If
    Secondary(when, 11) = Secondary(when, 11) + Recl(n, when)
    Secondary(when, 19) = Secondary(when, 6)
    Secondary(when, 22) = Secondary(when, 22) + Wrcp(n, when)
    Secondary(when, 23) = Secondary(when, 23) + Totl(n, when)
    Deprecy(n, MaxYr + 1) = Deprecy(n, MaxYr + 1) + Deprecy(n, when)
  Next n
Next when

adjust = 0
For when = 1 To maxprodyr - 1
  Secondary(maxprodyr, 21) = Secondary(maxprodyr, 21) - Secondary(when, 21)
  Secondary(when, 7) = (Sets(10) + adjust + Secondary(when, 23) - Secondary(when, 6)) * ((Sets(7) / 1000) * (Sets(8) / 100))
  Prop(when, 5) = Prop(when, 5) + ((Sets(10) + adjust + Secondary(when, 23) - Secondary(when, 6)) * ((Sets(7) / 1000) * (Sets(8) / 100)))
  adjust = adjust + Secondary(when, 23) - Secondary(when, 6)
Next when

End Sub

Public Sub autosalvage(n, salvtemp, maxprodyr, Salv() As Currency)

If CapitalData(n).Changed = False Then
  Salv(n, maxprodyr) = Salv(n, maxprodyr) + salvtemp
  CapitalData(n).SalvageAmount = Salv(n, maxprodyr)
  If CapitalData(n).SalvageAmount >= CapitalData(n).PurchaseAmount * 0.8439 Then
    CapitalData(n).SalvageAmount = CapitalData(n).PurchaseAmount * 0.8439
  End If
  CapitalData(n).SoldYear = maxprodyr
  If CapitalData(n).SoldYear < CapitalData(n).InvestYear And CapitalData(n).SoldYear > 0 Then
    CapitalData(n).SoldYear = CapitalData(n).InvestYear
  End If
End If

End Sub

Public Sub sensitivity()

End Sub
