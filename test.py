import matplotlib.pyplot as plt
import numpy as np

# Generate random batted ball data
num_balls = 100
theta = np.random.uniform(0, np.pi/2, num_balls)  # Random angle values (0 to 90 degrees)
r = np.random.uniform(0, 1, num_balls)  # Random radial distance values

# Create figure and axis
fig = plt.figure()
ax = fig.add_subplot(111, polar=True)

# Define the desired angle range (quarter of a circle)
angle_range = np.pi/2

# Split the field into 5 sections
num_sections = 5
section_angles = np.linspace(0, angle_range, num_sections + 1)

# Calculate the percentage of balls in each section
section_counts, _ = np.histogram(theta, bins=section_angles)
section_percentages = section_counts / num_balls

# Plot the field sections
bars = ax.bar(
    x=section_angles[:-1], 
    height=section_percentages, 
    width=section_angles[1] - section_angles[0], 
    alpha=0.5
)

# Remove grid lines and axis lines
ax.grid(False)
ax.spines['polar'].set_visible(False)

# Set plot properties
ax.set_xticklabels([])
ax.set_yticklabels([])
ax.set_title("Batted Ball Spray Chart")

# Display the plot
plt.show()
