# Creating PowerPoint slides using R, reticulate and pywin
# Author: Asif Salam
# email: asif.salam@yahoo.com
# Date: 2023-08-06

library(tidyverse)
library(stringr)
library(reticulate)

################# Create a dataset with filmography and film revenue #################
# download and read in the data files

# If downloading the files from the github repo
# download.file("https://raw.githubusercontent.com/asifsalam/r_and_powerpoint/main/data/clint_eastwood_films.tsv",
#              destfile = "clint_eastwood_films.tsv")
# download.file("https://raw.githubusercontent.com/asifsalam/r_and_powerpoint/main/data/clint_eastwood_box_office.csv", 
#              destfile = "clint_eastwood_box_office.csv")

# Output directory to save the created PowerPoint presentation
output_dir <- file.path(getwd(),"output")

clint_films <- read.table("./data/clint_eastwood_films.tsv",header=TRUE, stringsAsFactors=FALSE)
box_office <- read.table("./data/clint_eastwood_box_office.csv",header=TRUE, stringsAsFactors=FALSE)

# remove tv series
films <- clint_films %>% filter(!grepl("series",str_to_lower(additional_info)))
# Remove some films where the roles are uncredited - American Sniper, Casper, Breezy
films <- films[-which(films$key %in% c("americansniper","casper","breezy")),]

films <- left_join(films,box_office[,c("key","adjusted_gross")],by="key")
films$adjusted_gross[films$key=="crymacho"] <- 16510734
films$adjusted_gross[films$key=="themule"] <- 174800000
# clint_films_revenue$adjusted_gross[clint_films_revenue$key=="breezy"] <- 200000

film_revenue <- films %>% filter(!is.na(adjusted_gross)) %>% filter(adjusted_gross > 0) %>% arrange(desc(adjusted_gross))

nrow(films)


########## Creating the PowerPoint Slide ###################

actor_name <- "Clint Eastwood"

# Set up pywin
# Must install pywin32> pip install pywin32
pypath <- "C:/Program Files/Python310/"
use_python(pypath, required = T)
pywin <- import("win32com.client")

# Some prep
# Create the RGB function
pp_rgb <- function(r,g,b) {
    return(r + g*256 + b*256^2)
}

# Microsoft parameters
source("mso.txt")
# Some utility functions
source("utility_functions.R")

# Create the PowerPoint slide
pp = pywin$Dispatch('Powerpoint.Application')
pp[["Visible"]] = 1
presentation <- pp[["Presentations"]]$Add()

#slide1 <- presentation[["Slides"]]$Add(1,ms$ppLayoutBlank)
slide1 <- presentation[["Slides"]]$Add(1,ms$ppLayoutTitleOnly)
slide_width <- presentation[["PageSetup"]]$SlideWidth
slide_height <- presentation[["PageSetup"]]$SlideHeight

# Set some slide attributes
slide_color <- slide1[["ColorScheme"]]$Colors(ms$ppBackground)
slide_color[["RGB"]] <- pp_rgb(0,0,0)

# Create a background for the slide
img_file <- gsub("/","\\\\",paste(getwd(),"/","posters/clint_background_1.png",sep=""))
bg_image <- slide1[["Shapes"]]$AddPicture(img_file,TRUE,FALSE,0,0,slide_width,slide_height)
bg_rect <- slide1[["Shapes"]]$AddShape(ms$msoShapeRectangle,0, 0,slide_width, slide_height)
bg_rect_fill <- bg_rect[["Fill"]]
bg_rect_fill[["ForeColor"]][["RGB"]] <- pp_rgb(102, 25, 13)
bg_rect_fill[["Transparency"]] <- 0.1
bg_rect_line <- bg_rect[["Line"]]
bg_rect_line[["ForeColor"]][["RGB"]] <- pp_rgb(102,25,13)

slide_title <- slide1[["Shapes"]][["Title"]]
slide_title_color <- slide1[["ColorScheme"]]$Colors(ms$ppTitle)
slide_title_color[["RGB"]] <- pp_rgb(243,211,129)
slide_title$ZOrder(ms$msoBringToFront)

# Add a title
# AutoSize: https://msdn.microsoft.com/EN-US/library/office/ff745311(v=office.15).aspx
slide_title_frame <- slide_title[["TextFrame"]]
slide_title_frame[["AutoSize"]] <- ms$ppAutoSizeShapeToFitText
slide_title[["Top"]] <- 0
slide_title[["Left"]] <- 0
title_text <- slide_title[["TextFrame"]][["TextRange"]]
title_text[["Text"]] <- paste("Filmography: ",actor_name,sep="")
title_font <- title_text[["Font"]]
title_font[["Color"]][["RGB"]] <- pp_rgb(243,211,129)
title_font[["Size"]] <- 36
title_font[["Name"]] <- "Calibri"

# Add some decorative elements
diameter <- 100
# Add a line
line1 <- slide1[["Shapes"]]$AddLine(0,diameter/2,slide_width,diameter/2)
line1_attr <- line1[["Line"]]
line1_attr[["Weight"]] <- 1
line1_attr[["ForeColor"]][["RGB"]] <- pp_rgb(243,211,129)

# Add a circle, showing total box office earnings
circle1 <- slide1[["Shapes"]]$AddShape(ms$msoShapeOval,slide_width-diameter,0,diameter,diameter)
circle1[["Top"]] <- 0
circle1[["Left"]] <- slide_width - diameter
circle1[["Width"]] <- diameter
circle1[["Height"]] <- diameter
circle1_color <- circle1[["Fill"]]
circle1_color[["ForeColor"]][["RGB"]] <- pp_rgb(243,211,129)

# total_earnings <- format(sum(as.numeric(box_office$adjusted_gross))/1000000000,digits=3)
total_earnings <- format(sum(as.numeric(film_revenue$adjusted_gross))/1000000000,digits=3)
circle_frame <- circle1[["TextFrame"]]
circle_frame[["MarginTop"]] <- 0
circle_frame[["MarginLeft"]] <- 0
circle_frame[["MarginRight"]] <- 0
circle_frame[["MarginBottom"]] <- 0
circle_text <- circle1[["TextFrame"]][["TextRange"]]
circle_text[["Text"]] <- paste(total_earnings,"BUSD","")
circle_font <- circle_text[["Font"]]
circle_font[["Name"]] <- "Calibri"
circle_font[["Size"]] <- 24
circle_font[["Color"]][["RGB"]] <- pp_rgb(102,25,13)
circle_font[["Bold"]] <- 1
circle_line <- circle1[["Line"]]
circle_line[["Weight"]] <- 2
circle_line[["ForeColor"]][["RGB"]] <- pp_rgb(102,25,13)

earnings_text <- slide1[["Shapes"]]$AddTextbox(ms$msoTextOrientationHorizontal,
                                               slide_width-4*diameter,diameter/2-25,diameter*3+1,20)
earnings_range <- earnings_text[["TextFrame"]][["TextRange"]]
earnings_range[["Text"]] <- "Total Box Office Earnings"
earnings_font <- earnings_range[["Font"]]
earnings_font[["Color"]] <- pp_rgb(243,211,129)
earnings_font[["Size"]] <- 20
# When you are returning an object, you need to create a variable, and then set the properties
# This doesn't work
# earnings_range[["ParagraphFormat"]][["Alignment"]] <- ms$ppAlignRight
earnings_para <- earnings_range[["ParagraphFormat"]]
earnings_para[["Alignment"]] <- ms$ppAlignRight

# Place the poster images on the slide. 
# There are 60 movie images that need to be placed, so 20 columns by 3 rows
# films <- clint_films_revenue %>% arrange(id)
# films$id <- 1:nrow(films)
num_cols <- 20
num_rows <- ceiling(nrow(films)/num_cols)

# Based on the number of rows and columns, calculate the image height and width
image_width=floor((slide_width)/num_cols)

# We want to leave some space on the top for other things. The bottom half will contain the images.
image_offset <- 3
image_height=floor(slide_height/(num_rows+image_offset))

# We'll store a reference to the images in order to manipulate them later
images <- list()
image <- NULL

# Check how the images look when placed on the slide
for (i in 1:nrow(films)) {
    x = 0 + image_width * ((i-1) %% num_cols)
    y = image_height*image_offset + image_height * ((i-1) %/% num_cols)
    # Assuming that the images are in the working directory
    image_file <- gsub("/","\\\\",paste(getwd(),"/",films$poster_file[i],sep=""))
    images[[as.character(i)]] <- slide1[["Shapes"]]$AddPicture(image_file,TRUE,FALSE,x+1,y+1,image_width-2,image_height-2)
    image <- images[[as.character(i)]]
    
    line <- image[["Line"]]
    line[["Style"]] <- ms$msoLineDash
    line[["Weight"]] <- 2
    line[["ForeColor"]][["RGB"]] <- pp_rgb(243,211,129)
}

# delete_images(images)

# It's all a bit static right now. Add some animation.

# To animate the shapes we have created, we need a timeline sequence
# Add a sequence to the slide timeline: https://msdn.microsoft.com/en-us/library/office/ff746823.aspx
seq_main <- slide1[["TimeLine"]][["MainSequence"]]

# How animation works:
# - Three things: the object, an effect (fade, swivel etc.), and a timeline sequence in which it is to be placed.
# - Add an effect to a shape, on a timeline sequence, with a trigger (using seq.AddEffect, returns an Effect object)
# - The effect object is used to set the effect behaviours, path, trigger, timing and other parameters

animation_start(seq_main,slide_title,ms$msoAnimEffectDescend,ms$msoAnimTriggerWithPrevious,
                0, -20,0,0,1,0)
animation_start(seq_main,line1,ms$msoAnimEffectFly,ms$msoAnimTriggerAfterPrevious,
                100, diameter/slide_width,0,diameter/slide_height,1,0)
animation_start(seq_main,circle1,ms$msoAnimEffectFly,ms$msoAnimTriggerWithPrevious,
                -100, diameter/slide_height,0,diameter/slide_height,1,0)
animation_start(seq_main,earnings_text,ms$msoAnimEffectFly,ms$msoAnimTriggerWithPrevious,
                0, 100,(slide_width-4*diameter)/slide_width,diameter/slide_height,1,0)


# As far as I can tell, the way animation seems to works in PowerPoint is as follows:
# 1 - An animation effect (fade, swivel and so on) is applied to a specific object
# 2 - The animation is added to a timeline, which specifies the sequence in which the effect will be executed
# 3 - The animation can include how it is triggered, the duration, and some effect specific behaviours (such as the path)
# 4 - An external trigger can also be specified
# We can create a function that will apply animation to a shape, in this case the poster image
# The goal is to move the image from one point to another
# This function takes a timeline (sequence), the poster image that will be animated, 
# the button that will trigger the animation, the path along which the image will move
# and the duration, and applies the animation and parameters to the target poster image
# (See the section at the bottom for function - animate_shape())

# Add buttons which will sort the images by title or release year
# Add some explanatory text
sort_text <- slide1[["Shapes"]]$AddShape(ms$msoShapeRectangle,slide_width - 400,10,100,40)
sort_range <- sort_text[["TextFrame"]][["TextRange"]]
sort_para <- sort_range[["ParagraphFormat"]]
sort_para[["Alignment"]] <- ms$ppAlignLeft
sort_range[["Text"]] <- "Sort posters by: "
sort_font <- sort_range[["Font"]]
sort_font[["Size"]] <- 14
sort_font[["Color"]] <- pp_rgb(233,174,27)
sort_text[["Width"]] <- 110
sort_text[["Height"]] <- 15
sort_text[["Top"]] <- image_height*image_offset - 20 - 3
#sort_text[["Left"]] <- slide_width - 205
sort_text[["Left"]] <- 0
sort_fill <- sort_text[["Fill"]]
sort_fill[["Visible"]] <- 0
sort_line <- sort_text[["Line"]]
sort_line[["Visible"]] <- 0

# Create buttons that will sort and animate the poster images - alphanumeric
# btn2 <- slide1[["Shapes"]]$AddShape(ms$msoShapeRectangle,120,150,120,30)
button_alpha <- slide1[["Shapes"]]$AddShape(ms$msoShapeRectangle,120,150,120,30)
bta <- button_alpha[["TextFrame"]][["TextRange"]]
bta[["Text"]] <- "Title"

bta_line <- button_alpha[["Line"]]
bta_line[["ForeColor"]][["RGB"]] <- pp_rgb(243,211,129)

bta_fill <- button_alpha[["Fill"]]
bta_fill[["Visible"]] <- 1
bta_fill[["Transparency"]] <- 0.95
bta_rgb <- bta_fill[["ForeColor"]][["RGB"]]

bta_font <- bta[["Font"]]
bta_font[["Size"]] <- 14
bta_font[["Color"]] <- pp_rgb(243,211,129)

button_alpha[["Width"]] <- 90
button_alpha[["Height"]] <- 15
button_alpha[["Top"]] <- image_height*image_offset - 20 - 3
button_alpha[["Left"]] <- sort_text[["Width"]] - 2

# Create buttons that will sort and animate the poster images - date
button_date <- slide1[["Shapes"]]$AddShape(ms$msoShapeRectangle,200+160,100,150,30)
btd <- button_date[["TextFrame"]][["TextRange"]]
btd[["Text"]] <- "Release Year"

btd_fill <- button_date[["Fill"]]
btd_fill[["Visible"]] <- 1
btd_fill[["Transparency"]] <- 0.95

btd_line <- button_date[["Line"]]
btd_line[["ForeColor"]][["RGB"]] <- pp_rgb(243,211,129)

btd_font <- btd[["Font"]]
btd_font[["Size"]] <- 14
btd_font[["Color"]] <- pp_rgb(243,211,129)

button_date[["Width"]] <- 90
button_date[["Height"]] <- 15
button_date[["Top"]] <- image_height*image_offset - 20 - 3
button_date[["Left"]] <- sort_text[["Width"]] - 2 + button_alpha[["Width"]] + 20

# https://msdn.microsoft.com/EN-US/library/office/ff745511.aspx
animation_start(seq_main,sort_text,ms$msoAnimEffectWipe,ms$msoAnimTriggerAfterPrevious,
                0, 0,0,0,1,0)
animation_start(seq_main,button_alpha,ms$msoAnimEffectDissolve,ms$msoAnimTriggerWithPrevious,
                0, 0,0,0,1,0)
animation_start(seq_main,button_date,ms$msoAnimEffectDissolve,ms$msoAnimTriggerWithPrevious,
                0, 0,0,0,1,0)
#button_alpha$Delete()

# Add the interactive sequences, one per sort type
seq_alpha = slide1[["TimeLine"]][["InteractiveSequences"]]$Add()
seq_date = slide1[["TimeLine"]][["InteractiveSequences"]]$Add()


for (i in 1:nrow(films)) {
    x = 1 + image_width * ((i-1) %% num_cols)
    y = image_height*image_offset + image_height * ((i-1) %/% num_cols)
    image <- images[[as.character(i)]]

    # The position of this title, when sorted alphanumerically. That will dictate the positioning, and therefore the path to follow 
    index <- which(films$title[order(films$title)]==films$title[i]) - 1
    # Percentage of the slide_width the image must move in order to get to the new position
    l1 <- format((0 + image_width * (index %% num_cols) - x)/slide_width,digits=3)
    # Percentage of slide_height the image must move in order to get to the new position
    l2 <- format((image_height*image_offset + image_height * (index %/% num_cols) - y)/slide_height,digits=3)
    # Path - from current location (0,0) to new location (l1,l2)
    path <- paste0("M0,0 L",l1,",",l2)
    cat("i: ",i," - index: ",index ," - path: ",path,"\n")
    # Set the motion path to new location on alphanumeric sort
    animate_image(seq_alpha,image,button_alpha,path,2.0)
    # Set the motion path back to original location on release year sort
    path <- paste0("M",l1,",",l2," L0,0")
    animate_image(seq_date,image,button_date,path,2.0)
    # All images must move at the same time
    trigger_seq <- ms$msoAnimTriggerWithPrevious
    # Except for the first one
    if (i == 1) trigger_seq <- ms$msoAnimTriggerAfterPrevious
    
    # Set the entrance animation
    animation_start(seq_main,image,ms$msoAnimEffectDissolve,trigger_seq,
                    0, 0,0,0,0.5,0.1*i)
    
    # Link to the film's site, and create a tooltip
    link <- image$ActionSettings(ms$ppMouseClick)[["Hyperlink"]]
    link[["Address"]] <- films$film_url[i]
    link[["ScreenTip"]] <- paste0(films$title[i],"\nCharacter: ",films$character[i],"\nRelease Year: ",films$release_year[i])
}

# Better but still somewhat dissatifying. If the wrong sort button is clicked, 
# the images animate from the start of the path back to their current positions. Very tacky.

button_date$Delete()
button_alpha$Delete()
sort_text$Delete()

# We need a couple of Interactive sequences
seq_alpha2 = slide1[["TimeLine"]][["InteractiveSequences"]]$Add()
seq_date2 = slide1[["TimeLine"]][["InteractiveSequences"]]$Add()

# Recreate the text and buttons, on the same location
sort_text <- slide1[["Shapes"]]$AddShape(ms$msoShapeRectangle,slide_width - 400,10,100,40)
sort_range <- sort_text[["TextFrame"]][["TextRange"]]
sort_para <- sort_range[["ParagraphFormat"]]
sort_para[["Alignment"]] <- ms$ppAlignLeft
sort_range[["Text"]] <- "Sort posters by: "
sort_font <- sort_range[["Font"]]
sort_font[["Size"]] <- 14
sort_font[["Color"]] <- pp_rgb(233,174,27)
sort_text[["Width"]] <- 110
sort_text[["Height"]] <- 15
sort_text[["Top"]] <- image_height*image_offset - 20 - 3
#sort_text[["Left"]] <- slide_width - 205
sort_text[["Left"]] <- 0
sort_fill <- sort_text[["Fill"]]
sort_fill[["Visible"]] <- 0
sort_line <- sort_text[["Line"]]
sort_line[["Visible"]] <- 0

button_alpha <- slide1[["Shapes"]]$AddShape(ms$msoShapeRectangle,120,150,120,30)
bta <- button_alpha[["TextFrame"]][["TextRange"]]
bta[["Text"]] <- "Title"

bta_line <- button_alpha[["Line"]]
bta_line[["ForeColor"]][["RGB"]] <- pp_rgb(243,211,129)

bta_fill <- button_alpha[["Fill"]]
bta_fill[["Visible"]] <- 1
bta_fill[["Transparency"]] <- 0.95
bta_rgb <- bta_fill[["ForeColor"]][["RGB"]]

bta_font <- bta[["Font"]]
bta_font[["Size"]] <- 14
bta_font[["Color"]] <- pp_rgb(243,211,129)

button_alpha[["Width"]] <- 90
button_alpha[["Height"]] <- 15
button_alpha[["Top"]] <- image_height*image_offset - 20 - 3
button_alpha[["Left"]] <- sort_text[["Width"]] - 2

button_date <- slide1[["Shapes"]]$AddShape(ms$msoShapeRectangle,200+160,100,150,30)
btd <- button_date[["TextFrame"]][["TextRange"]]
btd[["Text"]] <- "Release Year"

btd_fill <- button_date[["Fill"]]
btd_fill[["Visible"]] <- 1
btd_fill[["Transparency"]] <- 0.95

btd_line <- button_date[["Line"]]
btd_line[["ForeColor"]][["RGB"]] <- pp_rgb(243,211,129)

btd_font <- btd[["Font"]]
btd_font[["Size"]] <- 14
btd_font[["Color"]] <- pp_rgb(243,211,129)

button_date[["Width"]] <- 90
button_date[["Height"]] <- 15
button_date[["Top"]] <- image_height*image_offset - 20 - 3
button_date[["Left"]] <- sort_text[["Width"]] - 2



# Entrance animation for the text and title button
animation_start(seq_main,sort_text,ms$msoAnimEffectWipe,ms$msoAnimTriggerAfterPrevious,
                0, 0,0,0,1,0)
animation_start(seq_main,button_alpha,ms$msoAnimEffectDissolve,ms$msoAnimTriggerWithPrevious,
                0, 0,0,0,1,0)

toggle_button(seq_alpha2,button_alpha,button_date,1)
toggle_button(seq_date2,button_date,button_alpha,1)

i = 2
for (i in 1:nrow(films)) {
    x = 1 + image_width * ((i-1) %% num_cols)
    y = image_height*image_offset + image_height * ((i-1) %/% num_cols)
    image <- images[[as.character(i)]]
    
    # The position of this title, when sorted alphanumerically. That will dictate the positioning, and therefore the path to follow 
    index <- which(films$title[order(films$title)]==films$title[i]) - 1
    # Percentage of the slide_width the image must move in order to get to the new position
    l1 <- format((0 + image_width * (index %% num_cols) - x)/slide_width,digits=3)
    # Percentage of slide_height the image must move in order to get to the new position
    l2 <- format((image_height*image_offset + image_height * (index %/% num_cols) - y)/slide_height,digits=3)
    # Path - from current location (0,0) to new location (l1,l2)
    path <- paste0("M0,0 L",l1,",",l2)
    cat("i: ",i," - index: ",index ," - path: ",path,"\n")
    # Set the motion path to new location on alphanumeric sort
    animate_image(seq_alpha,image,button_alpha,path,2.0)
    # Set the motion path back to original location on release year sort
    path <- paste0("M",l1,",",l2," L0,0")
    animate_image(seq_date,image,button_date,path,2.0)
    # Set the entrance animation
    animation_start(seq_main,image,ms$msoAnimEffectDissolve,trigger_seq,
                    0, 0,0,0,0.5,0.1*i)
    
    # Link to the film's site, and create a tooltip
    link <- image$ActionSettings(ms$ppMouseClick)[["Hyperlink"]]
    link[["Address"]] <- films$film_url[i]
    link[["ScreenTip"]] <- paste0(films$title[i],"\nCharacter: ",films$character[i],"\nRelease Year: ",films$release_year[i])
}

# Finally, save the file in the working directory
# Does not work: presentation$SaveAs(file.path(output_dir,"clint_eastwood_filmography"))
# Requires some path gymnastics to get this to work. The R path strings don't seem to work.
output_file <- gsub("/","\\\\",file.path(output_dir,"P3-A-Basic-Slide-Step-by-Step.pptx"))
presentation$SaveAs(output_file)
