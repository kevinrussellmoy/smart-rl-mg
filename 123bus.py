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
# ****************************************************
# * Examples Using the DSSText Object
# ****************************************************
# Load in an example circuit
DSSText.Command = r"Compile 'C:\Program Files\OpenDSS\IEEETestCases\123Bus\IEEE123Master.dss'"
# # Create a new capacitor
# DSSText.Command = "New Capacitor.C1 Bus1=1 Phases=3 kVAR=1200"
# DSSText.Command = "~ Enabled=false" # You can even use ~
# # Change the bus for Line L1
# DSSText.Command = "Line.L1.Bus1 = 5"
# # Export voltage to a csv file
# DSSText.Command = "Export Voltages"
# Filename = DSSText.Result
# print("File saved to: " + Filename)
# # ****************************************************
# # * Examples Using the DSSCircuit Object
# # ****************************************************
# # Step through every load and scale it up
iLoads = DSSCircuit.Loads.First
while iLoads:
    # Scale load by 10000%
    DSSCircuit.Loads.kW = DSSCircuit.Loads.kW * 1

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
# # ****************************************************
# # * Examples Using the DSSSolution Object
# # ****************************************************
# Solve the Circuit
DSSSolution.Solve()
if DSSSolution.Converged:
    print("The Circuit Solved Successfully")
else:
    print("The Circuit Did Not Solve Successfully")
# # Model effects of a large load pickup 30 seconds into a simulation
DSSText.Command = "New Monitor.Mon1 element=Load.S47 mode=0"
DSSSolution.StepSize = 1 # Set step size to 1 sec
DSSSolution.Number = 30 # Solve 30 seconds of the simulation
# Set the solution mode to duty cycle, which forces loads to use their
# "duty cycle" loadshape and allows time based simulation
DSSSolution.Mode = 6 # Code for duty cycle mode
DSSSolution.Solve()
DSSCircuit.Enable("Load.S47") # Enable the load
DSSSolution.Number = 30 # Solve another 30 seconds of simulation
DSSSolution.Solve()
print("Seconds Elapsed: " + str(DSSSolution.Seconds))
# Plot the voltage for the 60 seconds of simulation
DSSText.Command = "Plot monitor object=Mon1 Channels=(1,3,5)"

# Load in bus coordinates for plotting
DSSText.Command = r"Buscoords 'C:\Program Files\OpenDSS\IEEETestCases\123Bus\Buscoords.dat'"
# From IEEE123 test case in OpenDSS, "CircuitPlottingScripts.dss"
# These settings make a more interesting voltage plot since the voltages are generally OK for this case
# Voltages above     1.02            will be BLUE
# Voltages between   1.0 and 1.02    will be GREEN
# Voltages below     1.0             will be RED
# These are the default colors for the voltage plot
DSSText.Command = "Set markTransformers=yes"
DSSText.Command = "Set normvminpu=1.02"
DSSText.Command = "Set emergvminpu=1.0"
DSSText.Command = "Plot Circuit Voltage dots=y labels=y"

# Export voltage to a csv file
DSSText.Command = "Export Voltages"
Filename = DSSText.Result
print("File saved to: " + Filename)