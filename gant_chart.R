library(ggplot2)
library(readr)

#source https://onunicornsandgenes.blog/2020/03/07/using-r-simple-gantt-chart-with-ggplot2/

activities <- read_csv("activities.csv")

## Set factor level to order the activities on the plot
activities$activity <- factor(activities$activity,
                              levels = activities$activity[nrow(activities):1])

plot_gantt <- qplot(ymin = start,
                    ymax = end,
                    x = activity,
                    colour = category,
                    geom = "linerange",
                    data = activities,
                    size = I(5)) +
  scale_colour_manual(values = c("black", "grey", "purple", "yellow")) +
  coord_flip() +
  theme_bw() +
  theme(panel.grid = element_blank()) +
  xlab("") +
  ylab("") +
  ggtitle("Vacation planning")

