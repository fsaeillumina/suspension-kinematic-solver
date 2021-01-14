Option Explicit
Option Base 1



'------------------------------------------------------------
'This module uses trilateration to determine xyz coordinates for a
'   point given 3 hardpoints, and the distance from the hardpoints to
'   the unknown 4th point.
'
'Getting Started
'   In excel, open VBA, create a new module and copy in this code.
'   You will also need to copy the support.vb file into this 
'   module or another module in VBA. 
'   In the worksheet, you can use the function trilateration()
'
'How to use the function in excel
'   Explanation on arguments for the trilateration function is 
'   given at the end of this file. 
'
'   http://en.wikipedia.org/wiki/Trilateration
'       For trilateralization to work, it needs to be in a
'       certain coordinate system. See "Preliminary and final
'       computations" section of wikipedia.
'
'Please note: trilateralization can produce 1,2 or 0 results.
'   The last argument "pos12" will pick which of the 2 results
'   to return. For this reason, the spreadsheet checks to see
'   if any of the links have changed length, and if the CP is
'   the same at static.
'------------------------------------------------------------


Function ehatx( _
ByVal p1 As Variant, _
ByVal p2 As Variant, _
Optional ByVal p3 As Variant) _
As Variant
'returns variant/double array

Dim x(3) As Double
Dim temp As Variant

temp = SubVariant(p2, p1)

x(1) = CDbl(temp(1))
x(2) = CDbl(temp(2))
x(3) = CDbl(temp(3))

'Debug.Print p2(2) - p1(2)
'Debug.Print magnitude(x)

    ehatx = Array((p2(1) - p1(1)) / magnitude(x), _
(p2(2) - p1(2)) / magnitude(x), _
(p2(3) - p1(3)) / magnitude(x))
End Function

Function func_i( _
ByVal vP1 As Variant, _
ByVal vP3 As Variant, _
ByRef vEHatx As Variant) As Variant


func_i = dot_product(vEHatx, SubVariant(vP3, vP1))

End Function

Function ehaty( _
ByVal vP1 As Variant, _
ByVal vP3 As Variant, _
ByVal vI As Variant, _
ByVal vEHatx As Variant) _
As Variant

Dim dX(3) As Double 'this has nothing to do with integration (d/dx). it is a type double(d) and variable name x (X)
Dim vTemp As Variant

vEHatx(1) = vEHatx(1) * vI
vEHatx(2) = vEHatx(2) * vI
vEHatx(3) = vEHatx(3) * vI

vTemp = SubVariant(vP3, AddVariant(vP1, vEHatx)) 'P3-P1-vEhatx, note vEhatx equals (original_vehatx*i) at this point in procedure

dX(1) = CDbl(vTemp(1))
dX(2) = CDbl(vTemp(2))
dX(3) = CDbl(vTemp(3))

vTemp(1) = (vTemp(1)) / magnitude(dX) 'magnitude returns a double
vTemp(2) = (vTemp(2)) / magnitude(dX) 'magnitude returns a double
vTemp(3) = (vTemp(3)) / magnitude(dX) 'magnitude returns a double

ehaty = vTemp

End Function

Function Trilateration( _
ByVal vP1 As Variant, _
ByVal vP2 As Variant, _
ByVal vP3 As Variant, _
 _
ByVal vR1 As Variant, _
ByVal vR2 As Variant, _
ByVal vR3 As Variant, _
 _
Optional ByVal pos12 As Integer = 2) _
As Variant

'vP = [x,y,z] coordinates of known points
'vR = [d] distance from each known point to unknown hardpoint
'Note: distances must have same order as points.
'   ie. (first_point,second_point,third_point,distance_from_first_point,
'       distance_from_second_point, ...)
'TWO solutions are possible. Use pos12 to force a solution (0 or 1),
'   or let it default to one with a positive z value by blank=default=2

Dim vI As Variant
Dim vEHatx As Variant
Dim vEHaty As Variant
Dim vEHatz As Variant
Dim dX, dY, dZpos, dZneg, dZ As Double
Dim distance As Double 'distance from point 1 to point 2
Dim vnew As Variant
Dim dnew(3) As Double
Dim dJ As Double
Dim vTemp As Variant
Dim dTemp(3) As Double

vEHatx = ehatx(vP1, vP2, vP3) 'working
vI = func_i(vP1, vP3, vEHatx) 'working variant/double
vEHaty = ehaty(vP1, vP3, vI, vEHatx) ' working variant/double
vEHatz = cross_product(vEHatx, vEHaty)

'Format for magnitude
vTemp = SubVariant(vP2, vP1)
dTemp(1) = CDbl(vTemp(1))
dTemp(2) = CDbl(vTemp(2))
dTemp(3) = CDbl(vTemp(3))

distance = CDbl(magnitude(dTemp)) 'check this equals length between a and b
dJ = dot_product(vEHaty, SubVariant(vP3, vP1))

'-------------------------------------------------------------------------------------
'New coordinate system
'-------------------------------------------------------------------------------------

dX = vR1 ^ 2 - vR2 ^ 2 + distance ^ 2
dX = dX / (2 * distance)

dY = vR1 ^ 2 - vR3 ^ 2 + vI ^ 2 + dJ ^ 2
dY = dY / (2 * dJ)
dY = dY - vI * dX / dJ

If (vR1 ^ 2 - dX ^ 2 - dY ^ 2) < 0 Then
    Trilateration = "NoSoln"
    Exit Function
Else:   dZpos = Math.Sqr(vR1 ^ 2 - dX ^ 2 - dY ^ 2)
        dZneg = -Math.Sqr(vR1 ^ 2 - dX ^ 2 - dY ^ 2)
End If

If pos12 Then 'this includes default value of 2
        dZ = dZpos
Else:   dZ = dZneg
End If
'
'Debug.Print CDbl(vP1(1))
'Debug.Print dX * CDbl(vEHatx(1)) ' maybe I want vEHaty(2), z --> (3) ?
'Debug.Print dY * CDbl(vEHaty(1))
'Debug.Print dZpos * CDbl(vEHatz(1))
'Debug.Print

'Convert back to original coordinate system
'vnew(1) = CDbl(vP1(1)) + dX * CDbl(vEHatx(1)) + dY * CDbl(vEHaty(1)) + dZpos * CDbl(vEHatz(1))
'vnew(2) = CDbl(vP1(2)) + dX * CDbl(vEHatx(2)) + dY * CDbl(vEHaty(2)) + dZpos * CDbl(vEHatz(2))
'vnew(3) = CDbl(vP1(3)) + dX * CDbl(vEHatx(3)) + dY * CDbl(vEHaty(3)) + dZpos * CDbl(vEHatz(3))

'Find first values
dnew(2) = CDbl(vP1(2)) + dX * CDbl(vEHatx(2)) + dY * CDbl(vEHaty(2)) + dZ * CDbl(vEHatz(2))
dnew(3) = CDbl(vP1(3)) + dX * CDbl(vEHatx(3)) + dY * CDbl(vEHaty(3)) + dZ * CDbl(vEHatz(3))

If (pos12 = 2) And (dnew(2) < 0 Or dnew(3) < 0) Then 'These values are likely incorrect. Switch from zpos to zneg
dZ = dZneg
dnew(2) = CDbl(vP1(2)) + dX * CDbl(vEHatx(2)) + dY * CDbl(vEHaty(2)) + dZ * CDbl(vEHatz(2))
dnew(3) = CDbl(vP1(3)) + dX * CDbl(vEHatx(3)) + dY * CDbl(vEHaty(3)) + dZ * CDbl(vEHatz(3))
End If

dnew(1) = CDbl(vP1(1)) + dX * CDbl(vEHatx(1)) + dY * CDbl(vEHaty(1)) + dZ * CDbl(vEHatz(1))

Trilateration = dnew

End Function


'Sub AddDescriptionForFunctions()
'    Application.MacroOptions _
'        Macro:="trilateration", _
'        Description:="vP [x,y,z] coordinates of known points vR [d] distance from each known point to unknown hardpoint Note: distances must have same order as points. ie. (first_point,second_point,third_point,distance_from_first_point, distance_from_second_point,) TWO solutions are possible. Use pos12 to force a solution (0 or 1), or let it default to one with a positive z value by blank=default=2"
'        ArgumentDescriptions:=Array( _
'            "vP1: Matrix with three elements: [x,y,z] coordinates of first hardpoint", _
'            "vP2: Matrix with three elements: [x,y,z] coordinates of second hardpoint", _
'            "vP2: Matrix with three elements: [x,y,z] coordinates of third hardpoint", _
'            "vR1: Distance (1 number) from FIRST hardpoint (vP1) to hardpoint of unknown location", _
'            "vR1: Distance (1 number) from FIRST hardpoint (vP1) to hardpoint of unknown location", _
'            "vR1: Distance (1 number) from FIRST hardpoint (vP1) to hardpoint of unknown location", _
'            "pos12: (optional) Force the solver to pick a point. Leave blank to let the solver pick one that has a positive z value")
'
'End Sub

'Private Sub Workbook_Open()
'    Application.MacroOptions _
'        Macro:="trilateration", _
'        Description:="vP [x,y,z] coordinates of known points vR [d] distance from each known point to unknown hardpoint Note: distances must have same order as points. ie. (first_point,second_point,third_point,distance_from_first_point, distance_from_second_point,) TWO solutions are possible. Use pos12 to force a solution (0 or 1), or let it default to one with a positive z value by blank=default=2", _
'        ArgumentDescriptions:=Array( _
'            "vP1: Matrix with three elements: [x,y,z] coordinates of first hardpoint", _
'            "vP2: Matrix with three elements: [x,y,z] coordinates of second hardpoint", _
'            "vP2: Matrix with three elements: [x,y,z] coordinates of third hardpoint", _
'            "vR1: Distance (1 number) from FIRST hardpoint (vP1) to hardpoint of unknown location", _
'            "vR1: Distance (1 number) from FIRST hardpoint (vP1) to hardpoint of unknown location", _
'            "vR1: Distance (1 number) from FIRST hardpoint (vP1) to hardpoint of unknown location", _
'            "pos12: (optional) Force the solver to pick a point. Leave blank to let the solver pick one that has a positive z value")
'
'End Sub

