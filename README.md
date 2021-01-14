# suspension-kinematic-solver

'   
## Why is this useful?
Think of trilateration like a camera tripod. If we know the
xyz coordinates of the 3 feet of the tripod, and the length 
of the 3 legs, then we can mathematically determine the location 
of the camera. 

Given an input on one side of the suspension system (ie. moving 
the location of the tire contact patch up/down to simulate 
bump/roll), this method can determine the output on the other end 
(ie. damper displacement).  

The simplest method starts at the inboard push/pull rod spherical 
and solves for the xyz coordinates of the outboard spherical. 
