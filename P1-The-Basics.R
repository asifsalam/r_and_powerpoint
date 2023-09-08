# Scraping data from IMDB
# Author: Asif Salam
# email: asif.salam@yahoo.com
# Date: 2023-08-06


# Setup
# install.packages("reticulate")
library(reticulate)

# Define directory and file to save PowerPoint presentation
# Requires some path gymnastics to get this to work. The R path strings don't seem to work.
output_dir <- file.path(getwd(),"output")
output_file <- gsub("/","\\\\",file.path(output_dir,"PowerPoint_R_Basics.pptx"))

# Setup the python interface
pypath <- "C:/Program Files/Python310/"
use_python(pypath, required = T)

# pywin32 must be installed. use: >pip install pywin32

# load the Windows COM client from pywin32
# A basic tutorial on pywin32 https://pbpython.com/windows-com.html
# Additional info: http://timgolden.me.uk/python/win32_how_do_i/generate-a-static-com-proxy.html
pywin <- import("win32com.client")

# Some prep
# Create the RGb function
pp_rgb <- function(r,g,b) {
    return(r + g*256 + b*256^2)
}

# Microsft parameters
source("mso.txt")

# We can connect to the PowerPoint COM API
# Start up PowerPoint 
pp = pywin$Dispatch('Powerpoint.Application')

# Make the application visible
pp[["Visible"]] = 1
# pp$Visible <- TRUE

# Add a new presentation
presentation <- pp[["Presentations"]]$Add()
# Alternative way of adding a new presentation
# presentation <- pp$Presentations(1)

# The presentation is empty.  Add a slide to it.
slide1 <- presentation[["Slides"]]$Add(1,ms$ppLayoutBlank)
# or alternatively
# slide1 <- presentation$Slides$Add(1,ms$ppLayoutBlank)

# Add shapes and apply animation. 
# shp1 <- slide1$Shapes$AddShape(ms$msoShape12pointStar,20,20,100,100)
# https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shapes.addshape
# expression. AddShape( _Type_, _Left_, _Top_, _Width_, _Height_ )
# Add a 12 point star, 20 points from the top, and 20 points from the left, with a width and height of 100 points
shp1 <- slide1[["Shapes"]]$AddShape(ms$msoShape12pointStar,20,20,100,100)
shp2 <- slide1$Shapes$AddShape(ms$msoShapeHexagon,100,20,100,100)
shp3 <- slide1[["Shapes"]]$AddShape(ms$msoShapeCloud,180,20,100,200)

# Apply an animation effect to the shape, triggered immediately after the previous animation
# shp1 has 3 different animation effects applied, triggered one after another
# expression. AddEffect( _Shape_, _effectId_, _Level_, _trigger_, _Index_ )
slide1[["TimeLine"]][["MainSequence"]]$AddEffect(shp1,ms$msoAnimEffectFadedSwivel,
                                                 trigger=ms$msoAnimTriggerAfterPrevious)
slide1[["TimeLine"]][["MainSequence"]]$AddEffect(shp1,ms$msoAnimEffectPathBounceRight,
                                                 trigger=ms$msoAnimTriggerAfterPrevious)
slide1[["TimeLine"]][["MainSequence"]]$AddEffect(shp1,ms$msoAnimEffectSpin,
                                                 trigger=ms$msoAnimTriggerAfterPrevious)

# Store the animation effects from shp1
shp1$PickupAnimation()

# Apply the stored animation effects to the other shapes
shp2$ApplyAnimation()
shp3$ApplyAnimation()

# Add text to the shapes.  While this works, R files a complaint
shp1[["TextFrame"]][["TextRange"]][["Text"]] <- "Shp1-A"

# This way seems to function better
shp1_tr <- shp1[["TextFrame"]][["TextRange"]]
shp1_tr[["Text"]] <- "ONE"

# Set some shape attributes.  
# The `Fill` property is used for the colors, and the `Line` property for the border.
shp1_color <- shp1[["Fill"]]
shp1_color[["ForeColor"]][["RGB"]] <- (0+170*256+170*256^2)
# That's how the RGB value is calculated: r +  g*256 + b*256*256 

# Remove the line
shp1_line <- shp1[["Line"]]
shp1_line[["Visible"]] <- 0

# Now for the other shapes
# shp2$TextFrame$TextRange[["Text"]] <- "TWO"
shp2_tr <- shp2[["TextFrame"]][["TextRange"]]
shp2_tr[["Text"]] <- "TWO"
shp2_color <- shp2[["Fill"]]
shp2_color[["ForeColor"]][["RGB"]] <- pp_rgb(170,170,0)
# shp2$Line[["Visible"]] <- 1
shp2_line <- shp2[["Line"]]
shp2_line[["Visible"]] <- 0

shp3_tr <- shp3[["TextFrame"]][["TextRange"]]
shp3_tr[["Text"]] <- "THREE"
shp3_color <- shp3[["Fill"]]
shp3_color[["ForeColor"]][["RGB"]] <- pp_rgb(170,0,170)
shp3_line <- shp3[["Line"]]
shp3_line[["Visible"]] <- 0

# Finally, save the file in the output directory
presentation$SaveAs(output_file)

# To save in the working directory
# presentation$SaveAs(paste0(getwd(),"\\PowerPoint_R_P1_Basics.pptx"))