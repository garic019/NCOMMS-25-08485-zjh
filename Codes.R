## Figure1 ####
# Load and preprocess data
depression <- xlsx::read.xlsx("./depression.xlsx", sheetIndex = 1)

depression <- depression %>%
  mutate(mean = as.numeric(mean), lower = as.numeric(lower), upper = as.numeric(upper)) %>%
  mutate(Group = paste(Exposure, Model, Quartile, sep = "_"))

depression$Group <- factor(depression$Group,
                           levels = paste(
                             rep(c("Green space", "Blue space", "Natural environment"), each = 9),
                             rep(c("Model1", "Model2", "Model3"), times = 9),
                             rep(c("Q1", "Q2", "Q3"), each = 3, times = 3),
                             sep = "_"
                           )
)

# Save figure
pdf(file = "./Figures/Depression-300M.pdf", width = 10, height = 7.5, bg = "white")

# Plot hazard ratios and confidence intervals
dodge <- position_dodge(width = 0.7)

forest <- ggplot(depression, aes(x = Quartile, y = mean, ymin = lower, ymax = upper,
                                 color = Model, shape = Model)) +
  geom_hline(yintercept = 1, color = "darkgrey", linetype = "dashed", lwd = 0.8) +
  geom_pointrange(position = dodge, size = 0.9) +
  geom_errorbar(position = dodge, width = 0.3, lwd = 0.8) +
  facet_wrap(~Exposure, nrow = 1) +
  scale_color_manual(values = c("#666666", "#DFC286", "#608595")) +
  labs(y = "HR (95% CI)", x = "Quartile") +
  scale_y_continuous(breaks = seq(0.6, 1.2, 0.1), limits = c(0.6, 1.2), expand = c(0, 0)) +
  theme_classic() +
  theme(
    axis.text.x = element_text(size = 12),
    axis.text.y = element_text(size = 13),
    axis.title = element_text(size = 14, face = "bold"),
    legend.position = "top",
    legend.title = element_blank(),
    legend.text = element_text(size = 13),
    strip.text = element_text(size = 14, face = "bold"),
    plot.title = element_text(face = "bold", size = 15, margin = margin(10, 0, 10, 0)),
    axis.ticks.length = unit(0.3, "cm")
  )

print(forest)
dev.off()


## Figure2-3 ####
library(ggplot2)
library(tidyr)
library(dplyr)
library(openxlsx)

# Read in formatted dataset
data <- xlsx::read.xlsx("./forest_plot.xlsx", sheetIndex = 3)

data$Exposure <- factor(data$Exposure,
                        levels = c("Green space", "Blue space", "Natural environment"))
data <- data %>%
  mutate(mean = as.numeric(mean),lower = as.numeric(lower),upper = as.numeric(upper),
         Group = paste(Exposure,Air_pollution, Environment, sep = "_"))
# Order
data$Group <- factor(data$Group,
                     levels = paste(
                       rep(c("Green space", "Blue space", "Natural environment"), each = 9),   # Exposure
                       rep(c("Low", "Middle", "High"), times = 9),        # Model
                       rep(c("Tertile 1", "Tertile 2", "Tertile 3"),each = 3, times = 3),    # Quartile
                       sep = "_"))
pdf(file = "./Plot/PM10-300M.pdf",width = 10,height = 7.5,bg="white")
# Plot
dodge <- position_dodge(width = 0.7)
forest <- ggplot(data, aes(x = Air_pollution, y = mean, ymin = lower, ymax = upper,
                           color = Environment, shape = Environment)) +
  geom_hline(yintercept = 1, color = "darkgrey", linetype = "dashed", lwd = 0.8) +
  geom_pointrange(position = dodge, size = 0.7) + 
  geom_errorbar(position = dodge, width = 0, lwd = 0.8) +  
  scale_shape_manual(values = c(15, 15, 15, 15)) +  
  facet_wrap(~Exposure, nrow = 1) +  
  # scale_fill_manual(values=c("#666666", "#DFC286","#608595", "#ABC56E"))+
  scale_fill_manual(values=c("#546672", "#F5C75B","#308CC5", "#007B7A"))+
  scale_color_manual(values=c("#546672", "#F5C75B","#308CC5", "#007B7A"))+
  # scale_color_manual(values = c("#666666", "#608595", "#ABC56E")) +
  # scale_shape_manual(values = c(16, 17, 18)) +
  # Air pollution score / NO2 / NOx / PM2.5 / PM10
  labs(y = "HR (95% CI)", x = "Air pollution score") +
  scale_y_continuous(breaks = seq(0.2, 1.6, 0.3), limits = c(0.2, 1.6), expand = c(0, 0)) + 
  theme_classic() +
  theme(
    axis.text.x = element_text(size = 12),
    axis.text.y = element_text(size = 13),
    axis.title = element_text(size = 14, face = "bold"),
    legend.position = "top",
    legend.title = element_blank(),
    legend.text = element_text(size = 13),
    strip.text = element_text(size = 14, face = "bold"),
    plot.title = element_text(face = "bold", size = 15, margin = margin(10, 0, 10, 0)),
    axis.ticks.length = unit(0.3, "cm")
  )
forest
dev.off()



## Figure4-5 ####
library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)

# Read Excel data and label group/variable columns
data <- xlsx::read.xlsx("./forest_plot.xlsx", sheetIndex = 5)
data$Group <- ifelse(grepl("metabolic signature", data[,1], ignore.case = TRUE), data[,1], NA)
data$Group <- zoo::na.locf(data$Group)
data$Var <- ifelse(!grepl("metabolic signature", data[,1], ignore.case = TRUE), data[,1], NA)

# Filter Green & Nature data only
data <- data[1:12,]

# Reformat into tidy format for ggplot
clean_data <- data %>%
  filter(!is.na(Var)) %>%
  select(Group, Var,
         HR0 = HR, Low0 = CI95_low, High0 = CI95_high, P0 = P.value,
         HR1 = HR1, Low1 = CI95_low1, High1 = CI95_high1, P1 = P.value.1,
         HR2 = HR2, Low2 = CI95_low2, High2 = CI95_high2, P2 = P.value.2)

long_data <- bind_rows(
  clean_data %>% select(Group, Var, HR = HR0, Low = Low0, High = High0, P = P0) %>% mutate(Model = "Model 1"),
  clean_data %>% select(Group, Var, HR = HR1, Low = Low1, High = High1, P = P1) %>% mutate(Model = "Model 2"),
  clean_data %>% select(Group, Var, HR = HR2, Low = Low2, High = High2, P = P2) %>% mutate(Model = "Model 3")
)

# Format data types and factor levels
long_data <- long_data %>%
  filter(Var != "P for trend") %>%
  mutate(across(c(HR, Low, High), as.numeric)) %>%
  mutate(
    Var = factor(Var, levels = rev(unique(Var))),
    Model = factor(Model, levels = c("Model 3", "Model 2", "Model 1"))
  )

# Plot
p <- ggplot(long_data, aes(x = Var, y = HR, ymin = Low, ymax = High, color = Model, shape = Model)) +
  geom_hline(yintercept = 1, linetype = "dashed", color = "grey60") +
  geom_pointrange(position = position_dodge(width = 0.6), size = 0.9) +
  geom_errorbar(position = position_dodge(width = 0.6), width = 0.25, lwd = 0.8) +
  coord_flip() +
  facet_wrap(~Group, nrow = 1) +
  scale_color_manual(values = c("Model 1" = "#666666", "Model 2" = "#DFC286", "Model 3" = "#608595")) +
  scale_shape_manual(values = c("Model 1" = 16, "Model 2" = 17, "Model 3" = 18)) +
  scale_y_continuous(breaks = seq(0.75, 1.5, 0.25), limits = c(0.75, 1.5), expand = c(0, 0)) +
  theme_classic(base_size = 13) +
  theme(
    legend.position = "top",
    legend.title = element_blank(),
    strip.text = element_text(face = "bold", size = 13),
    axis.text.y = element_text(size = 12),
    plot.title = element_text(face = "bold", hjust = 0.5)
  ) +
  labs(
    y = "Hazard Ratio (95% CI)",
    x = NULL,
    title = "Forest Plot: Metabolic Signatures"
  ) +
  expand_limits(y = 1)

pdf(width = 12, height = 8, file = "./Figures/Green&Nature_metabolic.pdf")
print(p)
dev.off()
