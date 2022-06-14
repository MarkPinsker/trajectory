Function Contains(objCollection As Object, strName As String) As Boolean
    Dim o As Object
    On Error Resume Next
    Set o = objCollection(strName)
    Contains = (Err.Number = 0)
    Err.Clear
 End Function
Sub setupBalls()
' Setup lookup table of different types of balls with aerodynamic properties
Dim xlsBalls As Worksheet
Dim rowNo As Integer
Worksheets.Add.Name = "Balls"
Set xlsBalls = ThisWorkbook.Sheets("balls")
xlsBalls.Cells(1, 2) = "Diameter(m)"
xlsBalls.Cells(1, 3) = "Area(m2)"
xlsBalls.Cells(1, 4) = "Mass(Kg)"
xlsBalls.Cells(1, 5) = "CD"
For rowNo = 2 To 15
    xlsBalls.Cells(rowNo, 3).Formula = "=PI()*B" & Format(rowNo, "#") & "^2/4"
    xlsBalls.Cells(rowNo, 5) = 0.45
Next rowNo
xlsBalls.Cells(2, 1) = "42lb Cannon Ball"
xlsBalls.Cells(2, 2) = 0.17
xlsBalls.Cells(2, 4) = 19.091
xlsBalls.Cells(3, 1) = "6lb Cannon Ball"
xlsBalls.Cells(3, 2) = 0.087
xlsBalls.Cells(3, 4) = 2.727
xlsBalls.Cells(4, 1) = "Baseball"
xlsBalls.Cells(4, 2) = 0.074
xlsBalls.Cells(4, 4) = 0.145
xlsBalls.Cells(5, 1) = "Basketball"
xlsBalls.Cells(5, 2) = 0.243
xlsBalls.Cells(5, 4) = 0.62
xlsBalls.Cells(6, 1) = "Cricket ball"
xlsBalls.Cells(6, 2) = 0.072
xlsBalls.Cells(6, 4) = 0.16
xlsBalls.Cells(7, 1) = "Dodgeball"
xlsBalls.Cells(7, 2) = 0.216
xlsBalls.Cells(7, 4) = 0.454
xlsBalls.Cells(8, 1) = "Football"
xlsBalls.Cells(8, 2) = 0.22
xlsBalls.Cells(8, 4) = 0.43
xlsBalls.Cells(9, 1) = "Golf Ball"
xlsBalls.Cells(9, 2) = 0.043
xlsBalls.Cells(9, 4) = 0.046
xlsBalls.Cells(10, 1) = "Softball"
xlsBalls.Cells(10, 2) = 0.097
xlsBalls.Cells(10, 4) = 0.188
xlsBalls.Cells(11, 1) = "Table tennis ball"
xlsBalls.Cells(11, 2) = 0.04
xlsBalls.Cells(11, 4) = 0.003
xlsBalls.Cells(12, 1) = "Tennis ball"
xlsBalls.Cells(12, 2) = 0.067
xlsBalls.Cells(12, 4) = 0.058
xlsBalls.Cells(13, 1) = "Volleyball"
xlsBalls.Cells(13, 2) = 0.21
xlsBalls.Cells(13, 4) = 0.27
xlsBalls.Cells(14, 1) = ".22 lr"
xlsBalls.Cells(14, 2) = 0.0056
xlsBalls.Cells(14, 4) = 0.04
xlsBalls.Cells(14, 5) = 0.24
xlsBalls.Cells(15, 1) = ".458 Winchester Magnum"
xlsBalls.Cells(15, 2) = 0.01166
xlsBalls.Cells(15, 4) = 0.032
xlsBalls.Cells(15, 5) = 0.24
Set xlsBalls = Nothing
End Sub
Sub setupLocation()
' Setup lookup table of world locations with gravitational acceleration
Dim xlsLocation As Worksheet
 Worksheets.Add.Name = "Location"
Set xlsLocation = ThisWorkbook.Sheets("Location")
' More values can be obtained from https://en.wikipedia.org/wiki/Gravity_of_Earth
xlsLocation.Cells(1, 2) = "Gravity(ms-2)"
xlsLocation.Cells(1, 3) = "Altittude(m)"
xlsLocation.Cells(2, 1) = "London"
xlsLocation.Cells(2, 2) = 9.816
xlsLocation.Cells(2, 3) = 0
xlsLocation.Cells(3, 1) = "Singapore"
xlsLocation.Cells(3, 2) = 9.776
xlsLocation.Cells(3, 3) = 0
xlsLocation.Cells(4, 1) = "Anchorage"
xlsLocation.Cells(4, 2) = 9.826
xlsLocation.Cells(4, 3) = 0
Set xlsLocation = Nothing
End Sub

Sub setupTemperature()
' Setup lookup table of air temperatures and  their corresponding densities
Dim xlsTemperature As Worksheet
 Worksheets.Add.Name = "Temperature"
Set xlsTemperature = ThisWorkbook.Sheets("Temperature")

xlsTemperature.Cells(1, 1) = "temperature"
xlsTemperature.Cells(1, 2) = "rho(Kgm-3)"
xlsTemperature.Cells(2, 1) = 0
xlsTemperature.Cells(2, 2) = 1.2922
xlsTemperature.Cells(3, 1) = 5
xlsTemperature.Cells(3, 2) = 1.269
xlsTemperature.Cells(4, 1) = 10
xlsTemperature.Cells(4, 2) = 1.2466
xlsTemperature.Cells(5, 1) = 15
xlsTemperature.Cells(5, 2) = 1.225
xlsTemperature.Cells(6, 1) = 20
xlsTemperature.Cells(6, 2) = 1.2041
Set xlsTemperature = Nothing
End Sub
Sub setupGraphTypes()
' Setup lookup table of types of graph calculated by this macro
Dim xlsGraphTypes As Worksheet
Worksheets.Add.Name = "GraphTypes"
Set xlsGraphTypes = ThisWorkbook.Sheets("GraphTypes")
xlsGraphTypes.Cells(1, 1) = "Height"
xlsGraphTypes.Cells(1, 2) = "Distance(m)"
xlsGraphTypes.Cells(1, 3) = "Height(m)"

xlsGraphTypes.Cells(2, 1) = "Velocity"
xlsGraphTypes.Cells(2, 2) = "Time(s)"
xlsGraphTypes.Cells(2, 3) = "Velocity(m/s)"

xlsGraphTypes.Cells(3, 1) = "Height/time"
xlsGraphTypes.Cells(3, 2) = "Time(s)"
xlsGraphTypes.Cells(3, 3) = "Height(m)"

xlsGraphTypes.Cells(4, 1) = "xVelocity"
xlsGraphTypes.Cells(4, 2) = "Time(s)"
xlsGraphTypes.Cells(4, 3) = "Horizontal velocity(m/s)"

xlsGraphTypes.Cells(5, 1) = "yVelocity"
xlsGraphTypes.Cells(5, 2) = "Time(s)"
xlsGraphTypes.Cells(5, 3) = "Vertical velocity(m/s)"

xlsGraphTypes.Cells(6, 1) = "Angle of Elevation"
xlsGraphTypes.Cells(6, 2) = "Time(s)"
xlsGraphTypes.Cells(6, 3) = "Elevation angle(Degrees)"
Set xlsGraphTypes = Nothing
End Sub
Sub setupInitialConditions()
' Setup worksheet to calculate
Dim xlsInitialConditions As Worksheet
Dim iCol As Integer, noCols As Integer
Dim columnLetter  As String
Worksheets.Add.Name = "initial conditions"
Set xlsInitialConditions = ThisWorkbook.Sheets("initial conditions")
noCols = 4

xlsInitialConditions.Cells(1, 1) = "Ball"
xlsInitialConditions.Cells(2, 1) = "Diameter"
xlsInitialConditions.Cells(3, 1) = "Area"
xlsInitialConditions.Cells(4, 1) = "Mass"
xlsInitialConditions.Cells(5, 1) = "CD"
xlsInitialConditions.Cells(6, 1) = "Location"
xlsInitialConditions.Cells(7, 1) = "temperature"
xlsInitialConditions.Cells(8, 1) = "Air density"
xlsInitialConditions.Cells(9, 1) = "Gravity"
xlsInitialConditions.Cells(10, 1) = "Angle"
xlsInitialConditions.Cells(11, 1) = "Initial speed"
xlsInitialConditions.Cells(12, 1) = "Initial Height"
    
For iCol = 2 To noCols
    columnLetter = Split(xlsInitialConditions.Cells(1, iCol).Address, "$")(1)
    xlsInitialConditions.Cells(1, iCol) = "table tennis ball"
    xlsInitialConditions.Range(columnLetter & "1").Validation.Delete
    xlsInitialConditions.Range(columnLetter & "1").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=Balls!$A:$A"
    xlsInitialConditions.Range(columnLetter & "1").Validation.IgnoreBlank = True
    xlsInitialConditions.Range(columnLetter & "1").Validation.InCellDropdown = True
    xlsInitialConditions.Cells(2, iCol).Formula = "=VLOOKUP(" & columnLetter & "$1,Balls!$A:$B,2,FALSE)"
    xlsInitialConditions.Cells(3, iCol).Formula = "=PI()*" & columnLetter & "2^2/4"
    xlsInitialConditions.Cells(4, iCol).Formula = "=VLOOKUP(" & columnLetter & "1,Balls!$A:$D,4,0)"
    xlsInitialConditions.Cells(5, iCol).Formula = "=VLOOKUP(" & columnLetter & "1,Balls!$A:$E,5,FALSE)"
    xlsInitialConditions.Cells(6, iCol) = "Anchorage"
    xlsInitialConditions.Range(columnLetter & "6").Validation.Delete
    xlsInitialConditions.Range(columnLetter & "6").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=location!$A:$A"
    xlsInitialConditions.Range(columnLetter & "6").Validation.IgnoreBlank = True
    xlsInitialConditions.Range(columnLetter & "6").Validation.InCellDropdown = True
    xlsInitialConditions.Cells(7, iCol) = 20
    xlsInitialConditions.Range(columnLetter & "7").Validation.Delete
    xlsInitialConditions.Range(columnLetter & "7").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=temperature!$A:$A"
    xlsInitialConditions.Range(columnLetter & "7").Validation.IgnoreBlank = True
    xlsInitialConditions.Range(columnLetter & "7").Validation.InCellDropdown = True
    xlsInitialConditions.Cells(8, iCol).Formula = "=VLOOKUP(" & columnLetter & "7,temperature!$A:$B,2,FALSE)"
    xlsInitialConditions.Cells(9, iCol).Formula = "=VLOOKUP(" & columnLetter & "6,location!$A:$B,2,FALSE)"
    xlsInitialConditions.Cells(10, iCol) = 30
    xlsInitialConditions.Cells(11, iCol) = 5
    xlsInitialConditions.Cells(12, iCol) = 0
Next iCol
xlsInitialConditions.Cells(1, 3) = "tennis ball"
xlsInitialConditions.Cells(1, 4) = "football"

xlsInitialConditions.Cells(15, 1) = "Time delta"
xlsInitialConditions.Cells(15, 2) = 0.001

xlsInitialConditions.Cells(16, 1) = "Maximum time"
xlsInitialConditions.Cells(16, 2) = 40

xlsInitialConditions.Cells(17, 1) = "Divisions per metre"
xlsInitialConditions.Cells(17, 2) = 1000

xlsInitialConditions.Cells(18, 1) = "Type of graph"
xlsInitialConditions.Cells(18, 2) = "Velocity"
xlsInitialConditions.Range("B18").Validation.Delete
xlsInitialConditions.Range("B18").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=GraphTypes!$A:$A"
xlsInitialConditions.Range("B18").Validation.IgnoreBlank = True
xlsInitialConditions.Range("B18").Validation.InCellDropdown = True

xlsInitialConditions.Cells(19, 1) = "x axis label"
xlsInitialConditions.Cells(19, 2).Formula = "=VLOOKUP(B18,GraphTypes!$A:$B,2,FALSE)"

xlsInitialConditions.Cells(20, 1) = "y axis label"
xlsInitialConditions.Cells(20, 2).Formula = "=VLOOKUP(B18,GraphTypes!$A:$C,3,FALSE)"

xlsInitialConditions.Columns.AutoFit

End Sub
Sub mcrCalculate()
'
' mcrCalculate Macro
' Calculate trajectory of ball under gravity and air resistance
'
'
Dim graphType As String, graphTitle As String, rngChart As String, xAxislabel As String, yAxislabel As String
Dim ballTypeInTitle As Boolean, locationInTitle As Boolean, temperatureInTitle As Boolean
Dim initialAngleInTitle As Boolean, initialSpeedInTitle As Boolean, initialHeightInTitle As Boolean

    
Dim area(1000) As Double, mass(1000) As Double, cd(1000) As Double, airDensity(1000) As Double, gravity(1000) As Double
Dim initialHeight(1000) As Double
Dim initialSpeed(1000) As Double
Dim initialAngle(1000) As Double

Dim deltaT As Double, t As Double, timeLimit As Double, maxTime As Double
Dim angle(1000) As Double, alpha As Double
Dim k1 As Double, k2 As Double, k3 As Double, k4 As Double, k5 As Double, prevk5 As Double
Dim l1 As Double, l2 As Double, l3 As Double, l4 As Double, l5 As Double, prevl5 As Double
Dim rk1Speed As Double, rk2Speed As Double, rk3Speed As Double, rk4Speed As Double, rk5Speed As Double
Dim x As Double, y As Double, prevX As Double, prevY As Double
Dim maxX As Double, divisionsPerMetre As Double
Dim rowNo As Long, noRows As Long, colNo As Long
Dim iTrajectory As Integer, noTrajectories As Integer

Dim xlsInitialConditions As Worksheet, xlsLookup As Worksheet, xlsTrajectories As Worksheet
Dim chtGraph As Chart
Dim axsChart As Axis, axsTitChart As AxisTitle
Dim axsHorizontalChart As Axis, axsHorizontalTitChart As AxisTitle

Dim calculationStarted As Double
Const pi = 3.14159265358979
Const secondsInDay = 86400
Const ballTypeRowno = 1
Const locationRowno = 6
Const temperatureRowno = 7
Const initialAngleRowno = 10
Const initialSpeedRowno = 11
Const initialHeightRowno = 12
'======================================================================================================================
' Check workssheets are setup correctly
'======================================================================================================================
If Not Contains(Sheets, "Lookup") Then
    Worksheets.Add.Name = "Lookup"
End If
Set xlsLookup = ThisWorkbook.Sheets("Lookup")

If Not Contains(Sheets, "Trajectories") Then
    Worksheets.Add.Name = "Trajectories"
End If
Set xlsTrajectories = ThisWorkbook.Sheets("Trajectories")

If Not Contains(Sheets, "Location") Then
    Call setupLocation
End If

If Not Contains(Sheets, "Temperature") Then
    Call setupTemperature
End If

If Not Contains(Sheets, "Balls") Then
    Call setupBalls
End If

If Not Contains(Sheets, "GraphTypes") Then
    Call setupGraphTypes
End If

If Not Contains(Sheets, "initial conditions") Then
    Call setupInitialConditions
End If
Set xlsInitialConditions = ThisWorkbook.Sheets("initial conditions")
' Always rebuild chart from scratch
If Contains(Sheets, "Graph") Then
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Graph").Delete
    Application.DisplayAlerts = True
End If


xlsLookup.Cells.Clear
xlsTrajectories.Cells.Clear

calculationStarted = Now

deltaT = xlsInitialConditions.Cells(15, 2)
timeLimit = xlsInitialConditions.Cells(16, 2)
maxTime = 0

divisionsPerMetre = xlsInitialConditions.Cells(17, 2)
graphType = xlsInitialConditions.Cells(18, 2)
xAxislabel = xlsInitialConditions.Cells(19, 2)
yAxislabel = xlsInitialConditions.Cells(20, 2)



maxX = 0
'======================================================================================================================
' Set up array of starting values for each trajectory
'======================================================================================================================
noTrajectories = 1
ballTypeInTitle = True
locationInTitle = True
temperatureInTitle = True
initialAngleInTitle = True
initialSpeedInTitle = True
initialHeightInTitle = True
'======================================================================================================================
' Work out which parameters remain the same for all trajectories ( use them for titles )
' and those which differ ( use them for column headings)
'======================================================================================================================
While (xlsInitialConditions.Cells(ballTypeRowno, noTrajectories + 1) <> "")
    If noTrajectories > 1 Then
        If xlsInitialConditions.Cells(ballTypeRowno, noTrajectories) <> xlsInitialConditions.Cells(ballTypeRowno, noTrajectories + 1) Then ballTypeInTitle = False
        If xlsInitialConditions.Cells(locationRowno, noTrajectories) <> xlsInitialConditions.Cells(locationRowno, noTrajectories + 1) Then locationInTitle = False
        If xlsInitialConditions.Cells(temperatureRowno, noTrajectories) <> xlsInitialConditions.Cells(temperatureRowno, noTrajectories + 1) Then temperatureInTitle = False
        If xlsInitialConditions.Cells(initialAngleRowno, noTrajectories) <> xlsInitialConditions.Cells(initialAngleRowno, noTrajectories + 1) Then initialAngleInTitle = False
        If xlsInitialConditions.Cells(initialSpeedRowno, noTrajectories) <> xlsInitialConditions.Cells(initialSpeedRowno, noTrajectories + 1) Then initialSpeedInTitle = False
        If xlsInitialConditions.Cells(initialHeightRowno, noTrajectories) <> xlsInitialConditions.Cells(initialHeightRowno, noTrajectories + 1) Then initialHeightInTitle = False
    End If
    area(noTrajectories) = xlsInitialConditions.Cells(3, noTrajectories + 1)
    mass(noTrajectories) = xlsInitialConditions.Cells(4, noTrajectories + 1)
    cd(noTrajectories) = xlsInitialConditions.Cells(5, noTrajectories + 1)
    airDensity(noTrajectories) = xlsInitialConditions.Cells(8, noTrajectories + 1)
    gravity(noTrajectories) = xlsInitialConditions.Cells(9, noTrajectories + 1)
    initialAngle(noTrajectories) = xlsInitialConditions.Cells(initialAngleRowno, noTrajectories + 1)
    initialSpeed(noTrajectories) = xlsInitialConditions.Cells(initialSpeedRowno, noTrajectories + 1)
    initialHeight(noTrajectories) = xlsInitialConditions.Cells(initialHeightRowno, noTrajectories + 1)
    noTrajectories = noTrajectories + 1
Wend
noTrajectories = noTrajectories - 1

'======================================================================================================================
' Setup graph title
'======================================================================================================================
graphTitle = graphType
If (ballTypeInTitle) Then
    graphTitle = graphTitle & " " & xlsInitialConditions.Cells(ballTypeRowno, 2)
End If
If (locationInTitle) Then
    graphTitle = graphTitle & " " & xlsInitialConditions.Cells(locationRowno, 2)
End If
If (temperatureInTitle) Then
    graphTitle = graphTitle & " " & xlsInitialConditions.Cells(temperatureRowno, 2) & "°C"
End If
If (initialAngleInTitle) Then
    graphTitle = graphTitle & " elev. " & WorksheetFunction.Text(initialAngle(1), "General") & "°"
End If
If (initialSpeedInTitle) Then
    graphTitle = graphTitle & " Initially " & WorksheetFunction.Text(initialSpeed(1), "General") & " m/s"
End If
If (initialHeightInTitle) Then
    graphTitle = graphTitle & " Initial height " & WorksheetFunction.Text(initialHeight(1), "General") & " m"
End If


xlsTrajectories.Cells(1, 1) = "x"
'======================================================================================================================
' Loop through trajectories creating column headings and then calculating the values
'======================================================================================================================
For iTrajectory = 1 To noTrajectories
    t = 0
    x = 0
    y = initialHeight(iTrajectory)
'======================================================================================================================
' calculate elevation and initial speed for this trajectory
'======================================================================================================================
    k5 = initialSpeed(iTrajectory) * Cos(pi * initialAngle(iTrajectory) / 180)
    l5 = initialSpeed(iTrajectory) * Sin(pi * initialAngle(iTrajectory) / 180)

'======================================================================================================================
' Column headings
'======================================================================================================================
    rowNo = 1
    colNo = iTrajectory * 2 - 1
    columntitle = ""
    If (Not ballTypeInTitle) Then columntitle = xlsInitialConditions.Cells(ballTypeRowno, iTrajectory + 1) & " "
    If (Not locationInTitle) Then columntitle = columntitle & xlsInitialConditions.Cells(locationRowno, iTrajectory + 1) & " "
    If (Not temperatureInTitle) Then columntitle = columntitle & " " & xlsInitialConditions.Cells(temperatureRowno, iTrajectory + 1) & "°C "
    If (Not initialAngleInTitle) Then columntitle = columntitle & " elev. " & WorksheetFunction.Text(initialAngle(iTrajectory), "General") & "° "
    If (Not initialSpeedInTitle) Then columntitle = columntitle & " Initially " & WorksheetFunction.Text(initialSpeed(iTrajectory), "General") & " m/s "
    If (Not initialHeightInTitle) Then columntitle = columntitle & " Initial height " & WorksheetFunction.Text(initialHeight(iTrajectory), "General") & " m "
    If (columntitle = "") Then columntitle = graphType
    xlsTrajectories.Cells(rowNo, iTrajectory + 1) = columntitle
'======================================================================================================================
' Loop through time intervals to generate trajectory of projectile
'======================================================================================================================
    While (y >= 0 And t < timeLimit)
        t = t + deltaT
        alpha = airDensity(iTrajectory) * area(iTrajectory) * cd(iTrajectory) / mass(iTrajectory) / 2
    ' calculate new velocities using Runge-Kutta
        rk5Speed = (k5 ^ 2 + l5 ^ 2) ^ 0.5
        k1 = -deltaT * alpha * k5 * rk5Speed
        l1 = deltaT * (-gravity(iTrajectory) - alpha * l5 * rk5Speed)
        
        rk1Speed = ((k5 + k1 / 2) ^ 2 + (l5 + l1 / 2) ^ 2) ^ 0.5
        k2 = -deltaT * alpha * (k5 + k1 / 2) * rk1Speed
        l2 = deltaT * (-gravity(iTrajectory) - alpha * (l5 + l1 / 2) * rk1Speed)
        
        rk2Speed = ((k5 + k2 / 2) ^ 2 + (l5 + l2 / 2) ^ 2) ^ 0.5
        k3 = -deltaT * alpha * (k5 + k2 / 2) * rk2Speed
        l3 = deltaT * (-gravity(iTrajectory) - alpha * (l5 + l2 / 2) * rk2Speed)
        
        rk3Speed = ((k5 + k3) ^ 2 + (l5 + l3) ^ 2) ^ 0.5
        k4 = -deltaT * alpha * (k5 + k3) * rk3Speed
        l4 = deltaT * (-gravity(iTrajectory) - alpha * (l5 + l3) * rk3Speed)
        
        prevk5 = k5
        prevl5 = l5
        k5 = k5 + (k1 + 2 * k2 + 2 * k3 + k4) / 6
        l5 = l5 + (l1 + 2 * l2 + 2 * l3 + l4) / 6
        rk5Speed = (k5 ^ 2 + l5 ^ 2) ^ 0.5
' Calculate new displacements
        prevX = x
        prevY = y
        x = x + (k5 + prevk5) * deltaT / 2
        y = y + (l5 + prevl5) * deltaT / 2
        
' Output new point on trajectory

        rowNo = rowNo + 1
        
        If (y > 0) Then
            Select Case graphType
            Case "Height"
                xlsLookup.Cells(rowNo, colNo) = x
                xlsLookup.Cells(rowNo, colNo + 1) = y
            Case "Velocity"
                xlsLookup.Cells(rowNo, colNo) = t
                xlsLookup.Cells(rowNo, colNo + 1) = rk5Speed
            Case "xVelocity"
                xlsLookup.Cells(rowNo, colNo) = t
                xlsLookup.Cells(rowNo, colNo + 1) = k5
            Case "yVelocity"
                xlsLookup.Cells(rowNo, colNo) = t
                xlsLookup.Cells(rowNo, colNo + 1) = l5
            Case "Height/time"
                xlsLookup.Cells(rowNo, colNo) = t
                xlsLookup.Cells(rowNo, colNo + 1) = y
            Case "Angle of Elevation"
                xlsLookup.Cells(rowNo, colNo) = t
                If (k5 = 0) Then
                    xlsLookup.Cells(rowNo, colNo + 1) = 90
                Else
                    xlsLookup.Cells(rowNo, colNo + 1) = 180 * Atn(l5 / k5) / pi
                End If
            Case Else
                xlsLookup.Cells(rowNo, colNo) = "Error - incorrect graph type"
            End Select
' Set the displayed value to 0 if the ball has dropped below the ground
        Else
            xlsLookup.Cells(rowNo, colNo) = 0
            xlsLookup.Cells(rowNo, colNo + 1) = 0
        End If
            
            
        If (x > maxX) Then maxX = x
    Wend
    If t > maxTime Then maxTime = t
    
Next iTrajectory
'======================================================================================================================
' Create vlookups to convert x axis from time to range
'======================================================================================================================
Select Case graphType
Case "Height"
    noRows = maxX * divisionsPerMetre
Case "Height/time", "Velocity", "xVelocity", "yVelocity", "Angle of Elevation"
    noRows = maxTime / deltaT
End Select
If noRows > 1048575 Then noRows = 1048575

For rowNo = 1 To noRows
' First column is distance or time depending on graphtype
    Select Case graphType
    Case "Height"
        x = rowNo / divisionsPerMetre
        xlsTrajectories.Cells(rowNo + 1, 1) = x
    Case "Height/time", "Velocity", "xVelocity", "yVelocity", "Angle of Elevation"
        t = rowNo * deltaT
        xlsTrajectories.Cells(rowNo + 1, 1) = t
    End Select
    
    For iTrajectory = 1 To noTrajectories
        colNo = 2 * iTrajectory
' Convert column number to column letters
        columnLetter = Split(Cells(1, colNo - 1).Address, "$")(1)
        nextcolumnletter = Split(Cells(1, colNo).Address, "$")(1)
        xlsTrajectories.Cells(rowNo + 1, iTrajectory + 1).Formula = "=VLOOKUP(A" & Format(rowNo, "#") & ",Lookup!" & columnLetter & ":" & nextcolumnletter & ",2,TRUE)"
    Next iTrajectory
Next rowNo
'======================================================================================================================
' Create graph
'======================================================================================================================
Set chtGraph = ThisWorkbook.Sheets.Add(, Sheets(Sheets.Count), , xlChart)
chtGraph.Name = "Graph"
chtGraph.HasTitle = True
chtGraph.ChartTitle.Text = graphTitle


chtGraph.ChartType = xlXYScatter ' Needs to be -4169 not xlXYScatterlines=74

For iTrajectory = 1 To noTrajectories
    chtGraph.SeriesCollection.NewSeries
    chtGraph.SeriesCollection(iTrajectory).Name = xlsTrajectories.Cells(1, iTrajectory + 1)
    chtGraph.SeriesCollection(iTrajectory).XValues = xlsTrajectories.Range("A2:A" & Format(noRows, "#"))
    nextcolumnletter = Split(xlsTrajectories.Cells(1, iTrajectory + 1).Address, "$")(1)
    chtGraph.SeriesCollection(iTrajectory).Values = xlsTrajectories.Range(nextcolumnletter & "2:" & nextcolumnletter & Format(noRows, "#"))
    
    chtGraph.SeriesCollection(iTrajectory).Format.Line.Visible = msoTrue
    chtGraph.SeriesCollection(iTrajectory).Format.Line.Weight = 0.25
    chtGraph.SeriesCollection(iTrajectory).Format.Fill.Visible = msoFalse
    chtGraph.SeriesCollection(iTrajectory).MarkerStyle = xlMarkerStyleNone

Next iTrajectory

' Horizontal axis labels
Set axsHorizontalChart = chtGraph.Axes(1)
axsHorizontalChart.HasTitle = True
Set axsHorizontalTitChart = axsHorizontalChart.AxisTitle
axsHorizontalTitChart.Text = xAxislabel

' Vertical axis labels

Set axsChart = chtGraph.Axes(2)
axsChart.HasTitle = True
Set axsTitChart = axsChart.AxisTitle
axsTitChart.Text = yAxislabel


chtGraph.SetElement (msoElementLegendBottom)
chtGraph.SetElement (msoElementPrimaryCategoryGridLinesMajor)
chtGraph.SetElement (msoElementPrimaryValueGridLinesMinorMajor)
chtGraph.SetElement (msoElementPrimaryCategoryGridLinesMinorMajor)


MsgBox "Successful in " & Format(secondsInDay * (Now - calculationStarted), "#.0") & " seconds", , "Trajectory calculation"
End Sub

