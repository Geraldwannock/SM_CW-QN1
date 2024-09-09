import random
import pandas as pd
from openpyxl import Workbook

# Function to simulate the checkout process for n_customers over total_time
def simulate_checkout(n_customers, total_time):
    interarrival_times = [random.randint(1, 15) for _ in range(n_customers)]
    service_times = [random.randint(1, 8) for _ in range(n_customers)]
    
    arrival_times = [sum(interarrival_times[:i]) for i in range(1, n_customers + 1)]
    
    start_times = [0] * n_customers
    end_times = [0] * n_customers
    idle_times = [0] * n_customers
    total_idle_time = 0
    
    for i in range(n_customers):
        if i == 0:
            start_times[i] = arrival_times[i]
        else:
            start_times[i] = max(arrival_times[i], end_times[i - 1])
        
        end_times[i] = start_times[i] + service_times[i]
        
        if i > 0:
            idle_times[i] = start_times[i] - end_times[i - 1]
            total_idle_time += idle_times[i]
    
    time_in_system = [end_times[i] - arrival_times[i] for i in range(n_customers)]
    avg_time_in_system = sum(time_in_system) / n_customers
    
    idle_percentage = (total_idle_time / total_time) * 100
    
    return interarrival_times, service_times, arrival_times, start_times, end_times, idle_times, avg_time_in_system, idle_percentage

# Function to run the simulation and store the results in Excel
def run_simulation_to_excel(replications, n_customers, total_time, filename):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Simulation Results"
    
    # Write headers
    sheet.append(["Replication", "Customer", "Interarrival Time", "Service Time", "Arrival Time", "Start Time", "End Time", "Idle Time", "Average Time in System", "Idle Percentage"])
    
    for rep in range(replications):
        # Run simulation
        interarrival_times, service_times, arrival_times, start_times, end_times, idle_times, avg_time_in_system, idle_percentage = simulate_checkout(n_customers, total_time)
        
        for i in range(n_customers):
            sheet.append([rep+1, i+1, interarrival_times[i], service_times[i], arrival_times[i], start_times[i], end_times[i], idle_times[i], "", ""])
        
        # Add performance metrics at the end of each replication
        sheet.append([rep+1, "", "", "", "", "", "", "", avg_time_in_system, idle_percentage])
    
    # Save to Excel file
    workbook.save(filename)

# Parameters
n_customers = 20
total_time = 180
replications = 50
filename = "ecommerce_checkout_simulation.xlsx"  # Custom file name

# Run the simulation and generate the Excel file
run_simulation_to_excel(replications, n_customers, total_time, filename)

print(f"Simulation complete. Results saved to {filename}.")
