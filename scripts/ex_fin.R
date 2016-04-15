library( ReporteRs )
library( ggplot2 )
library( ggalt )
library( magrittr )
library( dplyr )
library( rtable )

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

# data -----
portofolio_perf = data.frame(
  port = rep( c("portofolio_1", "portofolio_2", "portofolio_3",
                   "portofolio_4", "portofolio_5", "portofolio_6"), each = 3 ),
  income = c( 36.2, 41, 43, 33.8, 37.8, 37.4, 34.8, 37.5, 33.3, 29.3,
              35.2, 31.7, 25, 27.3, 25.4, 24.2, 33, 24.5 ),
  year = rep(c(2013:2015), 6)
)

# plot -----
gg <- ggplot(portofolio_perf, aes( x = year, y = income, width = .75, fill = port )) +
  geom_bar(stat = "identity") + facet_wrap( ~ port, nrow = 1)  +
  mytheme +
  guides(fill = "none") +
  coord_cartesian(xlim = c(2012, 2016)) +
  labs(title = "Portofolio performances",
       subtitle = "years : 2013 to 2015",
       x = "Years", y = "Income",
       caption = paste0( "timestamp: ", format(Sys.time()) )
       )


options("ReporteRs-fontsize" = 10)

ft_table <- portofolio_perf %>% select( port, year, income) %>%
  rename(`Portofolio` = port, `Year` = year, `Performance` = income ) %>%
  FlexPivot(id = "Portofolio", transpose = "Year", columns = "Performance", columns.transpose = )
ft_table

pptx(template = "templates/template.pptx") %>%
  addSlide("Titre et contenu") %>%
  addTitle("Performances") %>%
  addFlexTable(ft_table, offx = .2, offy = 2, width = 3, height = 4) %>%
  addPlot(fun = print, x = gg, width = 6, height = 4, offx = 3.3, offy = 3, vector.graphic = TRUE, bg = "transparent") %>%
  writeDoc( "docs/example.pptx")

browseURL("docs/example.pptx")
