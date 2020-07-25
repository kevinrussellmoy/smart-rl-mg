import win32com.client
# ****************************************************
# * Initialize OpenDSS
# ****************************************************
# Instantiate the OpenDSS Object
try:
    DSSObj = win32com.client.Dispatch("OpenDSSEngine.DSS")
except:
    print ("Unable to start the OpenDSS Engine")
    raise SystemExit
# Set up the Text, Circuit, and Solution Interfaces
DSSText = DSSObj.Text
DSSCircuit = DSSObj.ActiveCircuit
DSSSolution = DSSCircuit.Solution

# Load in an example circuit
DSSText.Command = r"Compile 'C:\Program Files\OpenDSS\IEEETestCases\13Bus\IEEE13Nodeckt.dss'"

# Step through every load and scale it up
iLoads = DSSCircuit.Loads.First
while iLoads:
    # Scale load by 120%
    DSSCircuit.Loads.kW = DSSCircuit.Loads.kW * 2.0
    # Move to next load
    iLoads = DSSCircuit.Loads.Next
# # Set a capacitor's rated kVAR to 1200
# DSSCircuit.SetActiveElement("Capacitor.C83")
# DSSCircuit.ActiveDSSElement.Properties("kVAR").Val = 1200
# # Get bus voltages
BusNames = DSSCircuit.AllBusNames
Voltages = DSSCircuit.AllBusVmagPu

# # See what an arbitrary bus's voltage is
# print(BusNames[10] + "'s voltage mag in per unit is: " + str(Voltages[10]))

# Solve the Circuit
DSSSolution.Solve()
if DSSSolution.Converged:
    print("The Circuit Solved Successfully")
else:
    print("The Circuit Did Not Solve Successfully")


# Model effects of opening a line 30 seconds into a simulation
DSSText.Command = "New Monitor.Mon1 element=Line.692675 mode=0"
DSSSolution.StepSize = 1 # Set step size to 1 sec
DSSSolution.Number = 30 # Solve 30 seconds of the simulation
# Set the solution mode to duty cycle, forcing loads to use their "duty cycle" loadshape, allowing time based simulation
DSSSolution.Mode = 6 # Code for duty cycle mode

# Solve for initial condition (first 30 seconds)
DSSSolution.Solve()

# Disconnect load at bus 671 from circuit
DSSText.Command = "Open Load.671"  # not sure why this doesn't work with ' DSSCircuit.Disable("Load.671") ' ?

# Alternatively: Open line between bus 692 and bus 675
# DSSText.Command = "Open Line.684611"

# Solve another 30 seconds of simulation
DSSSolution.Number = 30
DSSSolution.Solve()

print("Seconds Elapsed: " + str(DSSSolution.Seconds))
# Plot the voltage for the 60 seconds of simulation
DSSText.Command = "Plot monitor object=Mon1 Channels=(1,3)"

# Line voltage should increase, as expected

# From IEEE123 test case in OpenDSS, "CircuitPlottingScripts.dss"
# These settings make a more interesting voltage plot since the voltages are generally OK for this case
# Voltages above     1.02            will be BLUE
# Voltages between   0.97 and 1.02   will be GREEN
# Voltages below     0.97            will be RED
# These are the default colors for the voltage plot
DSSText.Command = "Set markTransformers=yes"
DSSText.Command = "Set markSwitches=yes"
DSSText.Command = "Set normvminpu=1.02"
DSSText.Command = "Set emergvminpu=0.97"
DSSText.Command = "Plot circuit voltage dots=y labels=y"

# Export voltage to a csv file
DSSText.Command = "Export Voltages"
Filename = DSSText.Result
print("File saved to: " + Filename)