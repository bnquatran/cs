import numpy as np

# "1" indicates top part
# "2" indicates bottom part

units = "mm"
diameter2 = 8 # max fastener diameter or nominal diameter
tolerance1 = 0.15 # positional tolerance of clearance hole
tolerance2 = tolerance1 * 1.4 # positional tolerance of threaded hole
MaxClearanceHoleDepth1 = 13.5900
MinThreadDepth2 = 41.2500

H_min = diameter2 + tolerance1 + tolerance2 * (1 + 2 * MaxClearanceHoleDepth1 / MinThreadDepth2)
# print(str(H_min) + " " + units)

if units =="mm":
    H_min = H_min / 25.4

if H_min >= .001 and H_min <= .1:
    LowCostDrillToleranceUpper = .005
    LowCostDrillToleranceLower = -.001
elif H_min >= .101 and H_min <= .2:
    LowCostDrillToleranceUpper = .006
    LowCostDrillToleranceLower = -.001
elif H_min >= .201 and H_min <= .4:
    LowCostDrillToleranceUpper = .007
    LowCostDrillToleranceLower = -.002
elif H_min >= .401 and H_min <= .75:
    LowCostDrillToleranceUpper = .01
    LowCostDrillToleranceLower = -.005
elif H_min >= .751 and H_min <= 1:
    LowCostDrillToleranceUpper = .012
    LowCostDrillToleranceLower = -.006

# convert from asymmetric tolerance to symmetric tolerance
LowCostDrillTolerance = np.average([LowCostDrillToleranceUpper, abs(LowCostDrillToleranceLower)])
# print(str(LowCostDrillTolerance) + " in")

H_nom = H_min + abs(LowCostDrillToleranceLower)

# convert from asymmetric tolerance to symmetric tolerance
H_nom = np.average(H_nom + np.array([LowCostDrillToleranceUpper, LowCostDrillToleranceLower]))

print(str(H_nom) + " +/- " + str(LowCostDrillTolerance) + " in")
print(str(H_nom * 25.4) + " +/- " + str(LowCostDrillTolerance * 25.4) + " mm")

print("\n" + str(tolerance2) + " " + units)
