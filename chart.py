from matplotlib import pyplot as plt
from matplotlib import lines, transforms, text
import pandas as pd
from adjustText import adjust_text
import random

# Gets data from the excel file
pistolData = pd.read_excel('pistol.xlsx')

# Separates data into sub-sections
names = pistolData['Name']
damages = pistolData['Damage']
penetrations = pistolData['Penetration']
types = pistolData['Type']

# Array of colors that we want to use
possibleColors = ['red', 'darkorange', 'gold', 'lawngreen', 'deepskyblue', 'white', 'blueviolet', "teal", "palegreen"]

# Now we randomzie them
random.shuffle(possibleColors)

# Color vars
backgroundColor = '#292727'
axesColor = '#b5b0b0'

# Here is the array that will store the color values to use on the chart
colors = []
colorVal = 0

# Iterate to set each color value
for i in range(len(types)):
    if i != 0:
        # Double if to prevent index out of bounds
        if types[i] != types[i - 1]:
            colorVal += 1
    colors.append(possibleColors[colorVal])

# Tells the plot to use a dark theme
plt.style.use('dark_background')

# Sets the size of the image
figure = plt.figure(figsize=(16, 9), dpi=400)

# Change the background color
plt.rcParams['figure.facecolor'] = backgroundColor
plt.rcParams['axes.facecolor'] = backgroundColor

# This draws the grid, and sets the grid to be behind the other elements
plt.rc('axes', axisbelow=True)
plt.grid(color='#696363', linestyle=':', linewidth=1, zorder=0)

# Sets the numerical limits
plt.ylim(0, 60)
plt.xlim(30, 140)

# Where the armor classes lie on the penetration scale
armorClasses = [10, 20, 30, 40, 50, 60]

# Draws the lines for the armor classes
for i, armorClass in enumerate(armorClasses):
    plt.axhline(y=armorClass, color='#766f6f', zorder=1)
    plt.annotate("Class %i" % (i + 1), (plt.gca().get_xlim()[0] + 0.5, armorClass + 0.5), xytext=(plt.gca().get_xlim()[0] + 0.5, armorClass + 0.5))

# Draws the line for 1 shot in the chest
plt.axvline(x=80, color="#b5b0b0", linestyle="--", linewidth=1, zorder=1)
plt.annotate("Chest HP", (80.5, 0.5), xytext=(80.5, 0.5))

# Set the labels
plt.title("Pistol/SMG Rounds")
plt.ylabel("Penetration")
plt.xlabel("Damage")

# Makes scatter plot
plottedPoints = plt.scatter(damages, penetrations, c=colors, s=20, marker='d', zorder=3)

# Creates the labels for the data points
# DO NOT ADD ARROWS HERE, THAT WILL MAKE THE ADJUST TEXT FUNCTION MAKE IT SO THE ARROWS CAN'T CROSS AND IT BECOMES A HUGE MESS
annotations = [plt.annotate(name, (damages[i], penetrations[i]), va="bottom", ha="right", fontsize='x-small',  weight='bold', color=colors[i]) for (i, name) in enumerate(names)]

# Set the color of the axes
for ax in plt.gcf().get_axes():
    ax.tick_params(color=axesColor, labelcolor=axesColor)
    for spine in ax.spines.values():
        spine.set_edgecolor(axesColor)

# Now we adjust the text so it's not overlapping
adjust_text(annotations, arrowprops=dict(arrowstyle="-", color='w', lw=0.4), save_steps=False, force_points=(0.3, 0.6), force_text=(0.3, 0.6), expand_text=(2.4, 2.3), expand_points=(2.4, 2.3), va="bottom", ha="right")

# The adjust_text method makes new annotations specifically for the arrows, now we're going to go through them all and change their colors
for child in ax.get_children():
    # We're getting all objects, filtering out only annotations
    # The ones with no text will be the arrors
    if isinstance(child, text.Annotation) and child.get_text() == '':
        # Now we'll loop through all the annotations and find the matching one, and set the new color
        for annotation in annotations:
            if child.xy == annotation.xy:
                child.arrow_patch.set_color(colors[annotations.index(annotation)])

# Extra note: lines 90-97 took me several days, and made me question my sanity

# This makes the things to display in the legend
# lines.Line2D will make a circle (intuitive, I know) with the right color, but only for unique colors
handlers = [lines.Line2D(range(1), range(1), color=backgroundColor, marker='o', markersize=8, markerfacecolor=color) for (i, color) in enumerate(colors) if colors[i] != colors[i - 1]]

# Empty array for the types
arrayOfTypes = []

# This turns the pandas object into a list, there is to_list but that's not in the same order
for ammoType in types:
    # Only add unique types
    if ammoType not in arrayOfTypes:
        arrayOfTypes.append(ammoType)

# Here we add the legend, showing which color is to which ammo type
plt.legend(handlers, arrayOfTypes, fancybox=True, fontsize='small', facecolor=backgroundColor, framealpha=0)

# Saves the file and re-overrides the facecolor
plt.savefig('test.png', bbox_inches='tight', facecolor=backgroundColor)