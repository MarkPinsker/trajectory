# Purpose
This code is to calculate the trajectory of a projectile through the air assuming Newtonian resistance.
It generates Excel graphs of 6 different specified parameters such as trajectory, velocity against time or angle of elevation. 
![Sample](https://user-images.githubusercontent.com/107402485/173411195-b0fc2301-5ba9-4c0d-ae15-452db96618b3.PNG)

## Assumptions
1. Level ground.
2. No spin on the projectile.
3. No cross wind.
4. No viscous flow effects.

## Method
Calculation uses Runge-Kutta 4th order numerical integratation using specified time delta.

# To download the code
The macroChartV1 code can be loaded into an Excel spreadsheet by the following procedure:-

1. Open Excel spreadsheet.
2. Press Alt F11 to open a VBA window.
3. Insert -> module
4. Copy and paste in macroChartV1 code from github.
5. Close VBA window.

# To run the code
1. From Excel press Alt -> F8 to get a list of macros. 
2. Select mcrCalculate and click run. This will run the macro for three balls thrown from ground level at 5 m/s in Anchorage at 20C with an angle of elevation of 30 degrees.

## Add more projectiles 
1. Select cells  D1:D12 . Drag the square at the bottom right to E12 or further as required. 

## Remove projectiles
1. Select column E ( for example) right click and select delete.

## Change existing projectile type (in column E for example)
1. Click on cell E1 in worksheet "initial conditions" and select it from the dropdown.

## Change diameter of projectile (in column E for example)
1. Click on cell E2 in worksheet "initial conditions". This will override the value calculated for the projectile in cell E1, and automatically change the cross sectional area of the proctile in cell E3. 

## Change mass of projectile (in column E for example)
1. Click on cell E4 in worksheet "initial conditions" . This will override the value calculated for the projectile in cell E1.

## Change location (in column E for example)
1. Click on cell E6 in worksheet "initial conditions" and select it from the dropdown.

## Change air temperature (in column E for example)
1. Click on cell E7 in worksheet "initial conditions" and select it from the dropdown. Do not chnage it directly to a value not in the dropdown otherwise air density wil not be calculated and the calculation will not work!

## Change angle of projection (in column E for example)
1. Click on cell E10 in worksheet "initial conditions" . This value is metres.

## Change initial height (in column E for example)
1. Click on cell E12 in worksheet "initial conditions" . This value is in degrees from the horizontal.

## Change angle of projection (in column E for example)
1. Click on cell E10 in worksheet "initial conditions" . This value is in degrees from the horizontal.
## Add new projectile type (in column E for example)
1. Add it by copying "Balls" row 15 down to 16 and editing cells B16-E16 appropriately. Notice that C16 has a formula
2. Click on cell E1 in worksheet "initial conditions" and select it from the dropdown.
