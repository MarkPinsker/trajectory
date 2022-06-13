# Purpose
This code is to calculate the trajectory of a projectile through a fluid like air assuming Newtonian resistance.  
This just means that the drag or air resistance on the projectile is due to the momentum of the air impacting on it rather than the viscosity or stickiness of the air slowing it down. In most cases this is a good assumption but would not apply for a golf ball sinking through syrup.  
It generates Excel graphs of 6 different specified parameters such as trajectory, velocity against time or angle of elevation.  
Sample Excel chart:-  
![Sample](https://user-images.githubusercontent.com/107402485/173411195-b0fc2301-5ba9-4c0d-ae15-452db96618b3.PNG)

## Assumptions
1. Level ground.
2. No spin on the projectile( this can have a big effect on trajectory due to the Magnus effect).
3. No cross wind.
4. No viscous flow effects.
5. Subsonic air flow ( air resitance increases dramatically once past 330 m/s as it hits the sound barrier).

## Method
Calculation uses [Runge-Kutta](https://en.wikipedia.org/wiki/Runge%E2%80%93Kutta_methods) 4th order numerical integratation using specified time delta.

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
1. Select cells  D1:D12 . Drag the square at the bottom right to E12, F12 or further as required. 

## Remove projectiles
1. Select column D ( for example) right click and select delete.

## Change existing projectile type (in column D for example)
1. Click on cell D1 in worksheet "initial conditions" and select it from the dropdown.

## Change diameter of projectile 
1. Click on cell D2 in worksheet "initial conditions". This will override the value calculated for the projectile in cell D1, and automatically change the cross sectional area of the proctile in cell D3. 

## Change mass of projectile 
1. Click on cell D4 in worksheet "initial conditions" . This will override the value calculated for the projectile in cell D1.

## Change location
1. Click on cell D6 in worksheet "initial conditions" and select it from the dropdown. 
Notice that this will change the value of gravitational acceleration in cell D9.  
2. If the location is not in the dropdown then add a row to the "location" worksheet. More values are available in https://en.wikipedia.org/wiki/Gravity_of_Earth .

## Change air temperature
1. Click on cell D7 in worksheet "initial conditions" and select the temperature in Celsius from the dropdown. Do not change it directly to a value not in the dropdown otherwise air density will not be calculated and the calculation will not work!

## Change angle of projection
1. Click on cell D10 in worksheet "initial conditions" . This value is in degrees.

## Change initial speed
1. Click on cell D11 in worksheet "initial conditions" . This value is in metres per second.

## Change initial height
1. Click on cell D12 in worksheet "initial conditions" . This value is in degrees from the horizontal.

## Add new projectile type
1. Add it by copying "Balls" row 15 down to 16 and editing cells B16-E16 appropriately. Notice that C16 has a formula
2. Click on cell E1 in worksheet "initial conditions" and select it from the dropdown.
