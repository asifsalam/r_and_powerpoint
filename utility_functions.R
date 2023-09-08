# Utility functions for creating PowerPoint slides using R, reticulate and pywin
# Author: Asif Salam
# email: asif.salam@yahoo.com
# Date: 2023-08-06

delete_images <- function(images) {
    for (image in images) {
        image$Delete()
    }
}

toggle_button <- function(seq,button1,button2,duration=1.5) {
    # Enable exit effect on button1 when it is clicked
    # https://docs.microsoft.com/en-us/office/vba/api/powerpoint.effect
    effect <- seq$AddEffect(Shape=button1,effectId=ms$msoAnimEffectDissolve,
                            trigger=ms$msoAnimTriggerOnShapeClick)
    # https://docs.microsoft.com/en-us/office/vba/api/powerpoint.effect.exit
    effect[["Exit"]] <- 1
    effectTiming <- effect[["Timing"]]
    effectTiming[["TriggerType"]] <- ms$msoAnimTriggerOnShapeClick
    effectTiming[["TriggerShape"]] <- button1
    effectTiming[["Duration"]] <- duration
    
    # Disable the exit effect on button2 when button1 is clicked
    effect <- seq$AddEffect(Shape=button2,effectId=ms$msoAnimEffectDissolve,
                            trigger=ms$msoAnimTriggerOnShapeClick)
    effect[["Exit"]] <- 0
    effectTiming <- effect[["Timing"]]
    effectTiming[["TriggerType"]] <- ms$msoAnimTriggerWithPrevious
    effectTiming[["TriggerShape"]] <- button1
    effectTiming[["Duration"]] <- duration    
    
}

animation_start <- function(seq,shape,effectID,trigger,from_x,from_y,to_x,to_y,duration,delay_time) {
    
    effect <- seq$AddEffect(Shape=shape,effectId=effectID,trigger=trigger)
    ani <- effect[["Behaviors"]]$Add(ms$msoAnimTypeMotion)
    # MotionEffect Object: https://msdn.microsoft.com/EN-US/library/office/ff745317(v=office.15).aspx    
    aniMotionEffect <- ani[["MotionEffect"]]
    # https://msdn.microsoft.com/EN-US/library/office/ff745317.aspx
    aniMotionEffect[["FromX"]] <- from_x
    aniMotionEffect[["ToX"]] <- to_x
    aniMotionEffect[["FromY"]] <- from_y
    aniMotionEffect[["ToY"]] <- to_y   
    effectTiming <- effect[["Timing"]]
    effectTiming[["Duration"]] <- duration
    effectTiming[["TriggerDelayTime"]] <- delay_time
}


# Animate images (and bars) from one position to another
animate_image <- function(seq,image,trigger,path,duration=1.5) {
    effect <- seq$AddEffect(Shape=image,effectId=ms$msoAnimEffectPathDown,
                            trigger=ms$msoAnimTriggerOnShapeClick)
    ani <- effect[["Behaviors"]]$Add(ms$msoAnimTypeMotion)
    aniMotionEffect <- ani[["MotionEffect"]]
    aniMotionEffect[["Path"]] <- path
    effectTiming <- effect[["Timing"]]
    effectTiming[["TriggerType"]] <- ms$msoAnimTriggerWithPrevious
    effectTiming[["TriggerShape"]] <- trigger
    effectTiming[["Duration"]] <- duration
}

# Animate images on bar click
animate_bar_teeter <- function(seq,image,trigger,duration=1.5,delay_time=0) {
    effect <- seq$AddEffect(Shape=image,effectId=ms$msoAnimEffectTeeter,
                            trigger=ms$msoAnimTriggerOnShapeClick)
    effectTiming <- effect[["Timing"]]
    effectTiming[["TriggerType"]] <- ms$msoAnimTriggerWithPrevious
    effectTiming[["TriggerShape"]] <- trigger
    effectTiming[["Duration"]] <- duration
    effectTiming[["TriggerDelayTime"]] <- delay_time
    
}

animate_bar <- function(seq,image,trigger,duration=1.5,delay_time=0) {
    effect <- seq$AddEffect(Shape=image,effectId=ms$msoAnimEffectSpinner,
                            trigger=ms$msoAnimTriggerOnShapeClick)
    effectTiming <- effect[["Timing"]]
    effectTiming[["TriggerType"]] <- ms$msoAnimTriggerWithPrevious
    effectTiming[["TriggerShape"]] <- trigger
    effectTiming[["Duration"]] <- duration
    effectTiming[["TriggerDelayTime"]] <- delay_time
    
}


pp_rgb <- function(r,g,b) {
    return( r + g*256 + b*256^2)
}

scale_bar_height <- function(max_value,max_height,value) {
    bar_height <- value*max_height/max_value
}
