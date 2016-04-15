# load library ------------
library(dplyr)
library(tidyr)
library(ggplot2)
library(readr)
library(readxl)
library(nlme)
library(lsmeans)
library(rtable)

# gg theme ------------

mytheme <- theme(axis.ticks = element_line(colour = "gray80"),
  panel.grid.major = element_line(colour = "gray80"),
  panel.grid.minor = element_line(linetype = "blank"),
  axis.text.x = element_text(colour = "gray20"),
  axis.text.y = element_text(colour = "gray20"),
  panel.background = element_rect(fill = NA),
  legend.key = element_rect(fill = NA),
  legend.title = element_text(size = 12, face = "italic"),
  legend.background = element_rect(fill = NA))


# import dataset --------
# le dataset des mesures
data <- read_excel("data/dataset.xlsx")

# les baselines mesurées avant les essais - la mesure d'intensite avant traitement
baselines <- read_csv("data/baselines.csv")

# la `demographie`
demo <- read_csv("data/demo.csv")

# integration des baselines
datamart <- data %>%
  inner_join(baselines)

# get time as a number
datamart <- datamart %>%
  tidyr::extract(time, "time_value", "day_([[:digit:]]+)", remove = FALSE ) %>%
  mutate( time_value = as.integer(time_value) )

# estimate effects with lme and lsmeans --------
fit <- lme( fixed = value ~ baseline + treatment + time + area,
            random = ~ 1 | patient_id,
            data = datamart )

# extraction des LS Means
lsmeans_ <- lsmeans(fit, ~ time * treatment)


# create reporting elements ----
options("ReporteRs-fontsize" = 9)

mytable <- table(demo$Gender, demo$Type)
tfreq <- freqtable(mytable)

# summarize data for plotting
scale_x <- datamart %>%
  group_by( treatment, time_value ) %>%
  summarise(q25 = quantile(value, probs = .25 ),
            q50 = quantile(value, probs = .5 ),
            q75 = quantile(value, probs = .75 ))

# plot with ggplot2
ggsummary <- ggplot( scale_x, aes(x = time_value, y = q50)) +
  geom_point(size = .75 ) +
  geom_errorbar(aes(ymax = q75, ymin = q25) ) +
  facet_wrap(~ treatment, ncol = 1 ) +
  mytheme


# A FlexTable to show statistics in a nice (non tidy) table
summary_ <- datamart %>%
  group_by(treatment, time_value) %>%
  summarise(avg = sprintf( "%.2f", mean(value) ),
            sd = sprintf( "%.2f", sd(value) ),
            n = length(value) ) %>%
  FlexPivot(id = "time_value",
            transpose = "treatment", space.table = TRUE,
            columns = c("avg", "sd", "n") )




# extract lsmeans and format them
lsm_estimates <- summary(lsmeans_) %>%
  mutate(ci = sprintf("[%.2f - %.2f]", lower.CL, upper.CL),
         est. = sprintf("%.2f (%.2f)", lsmean, SE) ) %>%
  tidyr::extract(time, "time_value", "day_([[:digit:]]+)", remove = FALSE ) %>%
  mutate( time_value = as.integer(time_value) )

ft_est <- FlexPivot(lsm_estimates, id = "time",
  transpose = "treatment", space.table = TRUE,
  columns = c("est.", "ci") )

# plot estimates ---
myplot <- ggplot(data = lsm_estimates,
       aes(x = time_value, y = lsmean,
           color = treatment, group = treatment)) +
  geom_line(position = position_dodge(width = .5)) +
  scale_color_manual(values = c("#2779A7", "#000000", "#EC3939", "#45A38D")) +
  coord_cartesian(xlim = c(0, max(lsm_estimates$time_value) ) ) +
  labs(title = "Least-squares means for time and treatment",
       subtitle = sprintf("Log-restricted-likelihood: %.1f", logLik(fit) ),
       x = "Days",
       y = "Estimates for xxx",
       caption = "Linear mixed-effects model fit by REML" )

myplot <- myplot + mytheme


# report generation --------
docx(template = "templates/template.docx") %>%

  addTOC() %>%
  addPageBreak() %>%

  addTitle("Summary") %>%

  addTitle("Statistical summary", level = 2 ) %>%
  addFlexTable(summary_) %>%
  addParagraph("statistical summary", stylename = "rTableLegend") %>%
  addPageBreak() %>%
  addPlot(fun = print, x = ggsummary, width = 4, height = 7, vector.graphic = TRUE) %>%
  addParagraph("graphical summary", stylename = "rPlotLegend") %>%
  addPageBreak() %>%

  addTitle("Demography", level = 3 ) %>%
  addFlexTable(tfreq) %>%
  addParagraph("demo frequency table", stylename = "rTableLegend") %>%

  addPageBreak() %>%
  addTitle("Model") %>%

  addTitle("Quality", level = 2) %>%
  addFlexTable(FlexTable_quality(fit)) %>%

  addTitle("Coefficients", level = 2) %>%
  addFlexTable(FlexTable_coef(fit)) %>%

  addSection( landscape = TRUE ) %>%
  addTitle("Estimates", level = 2) %>%
  addPlot(fun = print, x = myplot, width = 7, height = 5, vector.graphic = TRUE) %>%
  addParagraph("LS means values", stylename = "rPlotLegend") %>%
  addPageBreak() %>%
  addFlexTable(ft_est) %>%
  addParagraph("LS means values", stylename = "rTableLegend") %>%
  addSection( landscape = FALSE ) %>%

  addTitle("List of tables", level = 1) %>%
  addTOC(stylename = "rTableLegend") %>%
  addTitle("List of graphics", level = 1) %>%
  addTOC(stylename = "rPlotLegend") %>%

  writeDoc( "docs/example.docx")

browseURL("docs/example.docx")
