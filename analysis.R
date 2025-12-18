# Load libraries

library(dplyr)
library(flextable)
library(officer)
library(xtable)

# --- Load Dataset ---
data <- thesisdata  # replace with actual data frame if different

# --- Fix: Clean driver type groupings if not yet done ---
data$type_of_driver <- trimws(data$type_of_driver)

# --- Recode Continuous Variables into Categories ---
data$age_group <- cut(data$age,
                      breaks = c(21, 29, 39, 49, 59, 69),
                      labels = c("22–29", "30–39", "40–49", "50–59", "60–69"),
                      right = TRUE)

data$driving_experience_group <- cut(data$years_of_driving_experience,
                                     breaks = c(0, 5, 10, 20, 30, 40),
                                     labels = c("1–5 yrs", "6–10 yrs", "11–20 yrs", "21–30 yrs", "31–40 yrs"),
                                     right = TRUE)

data$km_per_day_group <- cut(data$how_many_kilometers_do_you_cover_daily,
                             breaks = c(0, 50, 100, 200, 300, 400),
                             labels = c("≤50 km", "51–100 km", "101–200 km", "201–300 km", "301–370 km"),
                             right = TRUE)

# --- Categorize EV Ownership Duration (for EV drivers only) ---
data$ev_usage_duration_cat <- cut(data$how_long_have_you_owned_or_used_your_ev,
                                  breaks = c(0, 1, 2, 4, 6, Inf),
                                  labels = c("Less than 1 year", "1–2 years", "3–4 years", "5–6 years", "Over 6 years"),
                                  right = TRUE)

# --- Function to Generate Frequency Tables with Totals ---
generate_table <- function(var) {
  tbl <- table(var, useNA = "ifany")
  percent <- prop.table(tbl) * 100
  cum_percent <- cumsum(percent)
  
  result <- data.frame(
    Category = as.character(names(tbl)),
    Frequency = as.numeric(tbl),
    `Valid Percent` = round(as.numeric(percent), 2),
    `Cumulative Percent` = round(as.numeric(cum_percent), 2),
    stringsAsFactors = FALSE
  )
  
  total_row <- data.frame(
    Category = "Total",
    Frequency = sum(result$Frequency, na.rm = TRUE),
    `Valid Percent` = 100.00,
    `Cumulative Percent` = NA_real_,
    stringsAsFactors = FALSE
  )
  
  result <- rbind(result, total_row)
  return(result)
}

# --- Define Variables (Universal, ICE, EV) ---
main_vars <- data[c(
  "location_within_accra_metropolis", "sex", "marital_status", "age_group",
  "driving_experience_group", "monthly_income", "level_of_education_completed",
  "type_of_vehicle_currently_own_or_driven", "do_you_own_the_vehicle_you_drive_currently",
  "km_per_day_group", "do_you_support_the_conversion_or_retrofitting_of_conventional_vehicles_into_electric_vehicles",
  "do_you_think_government_is_doing_enough_to_support_electric_vehicle_adoption", "type_of_driver"
)]

ice_data <- subset(data, grepl("Internal Combustion Engine", type_of_driver, ignore.case = TRUE))
ice_vars <- ice_data[c(
  "have_you_heard_of_cars_that_run_solely_on_battery_power_without_a_fuel_engine",
  "are_you_likely_to_buy_another_car_in_the_next_20_years_or_sooner",
  "do_you_have_plans_of_buying_an_electric_vehicles_in_the_next_20_years_or_sooner"
)]

ev_data <- subset(data, grepl("Electric Vehicle", type_of_driver, ignore.case = TRUE))
ev_vars <- ev_data[c(
  "what_type_of_ev_do_you_drive",
  "ev_usage_duration_cat"
)]

# --- Create Word Document ---
doc <- read_docx()

doc <- body_add_par(doc, "Descriptive Statistics – Sociodemographic Data", style = "heading 1")

# --- Universal Variables ---
doc <- body_add_par(doc, "Universal Sociodemographic Variables", style = "heading 2")
for (var_name in names(main_vars)) {
  tbl <- generate_table(main_vars[[var_name]])
  ft <- flextable(tbl)
  ft <- autofit(ft)
  ft <- bold(ft, i = nrow(tbl), bold = TRUE, part = "body")
  ft <- set_caption(ft, caption = paste("Distribution of", gsub("_", " ", var_name)))
  doc <- body_add_flextable(doc, ft)
  doc <- body_add_par(doc, "", style = "Normal")
}

# --- ICE Drivers Only ---
doc <- body_add_par(doc, "ICE Drivers Only", style = "heading 2")
for (var_name in names(ice_vars)) {
  tbl <- generate_table(ice_vars[[var_name]])
  ft <- flextable(tbl)
  ft <- autofit(ft)
  ft <- bold(ft, i = nrow(tbl), bold = TRUE, part = "body")
  ft <- set_caption(ft, caption = paste("ICE Driver:", gsub("_", " ", var_name)))
  doc <- body_add_flextable(doc, ft)
  doc <- body_add_par(doc, "", style = "Normal")
}

# --- EV Drivers Only ---
doc <- body_add_par(doc, "EV Drivers Only", style = "heading 2")
for (var_name in names(ev_vars)) {
  tbl <- generate_table(ev_vars[[var_name]])
  ft <- flextable(tbl)
  ft <- autofit(ft)
  ft <- bold(ft, i = nrow(tbl), bold = TRUE, part = "body")
  ft <- set_caption(ft, caption = paste("EV Driver:", gsub("_", " ", var_name)))
  doc <- body_add_flextable(doc, ft)
  doc <- body_add_par(doc, "", style = "Normal")
}

# --- Save the Word Document ---
print(doc, target = "Sociodemographic_Tables.docx")



#Cross tabulation

library(tibble)

# Create simplified driver type group variable
data$driver_type_group <- dplyr::case_when(
  data$type_of_driver == "Internal Combustion Engine Vehicle driver (Patrol, diesel etc. engine vehicle)" ~ "ICE",
  data$type_of_driver == "Electric Vehicle Driver" ~ "EV",
  TRUE ~ NA_character_
)


# --- Define list of sociodemographic variables to cross-tabulate ---
cross_vars <- list(
  Location = data$location_within_accra_metropolis,
  Sex = data$sex,
  Marital_Status = data$marital_status,
  Age = data$age_group,
  Driving_Experience = data$driving_experience_group,
  Monthly_Income = data$monthly_income,
  Education = data$level_of_education_completed,
  Vehicle_Type = data$type_of_vehicle_currently_own_or_driven,
  Own_Vehicle = data$do_you_own_the_vehicle_you_drive_currently,
  Daily_Kilometers = data$km_per_day_group,
  Support_Retrofit = data$do_you_support_the_conversion_or_retrofitting_of_conventional_vehicles_into_electric_vehicles,
  Gov_Support = data$do_you_think_government_is_doing_enough_to_support_electric_vehicle_adoption
)

# Start Word document
doc <- read_docx()
doc <- body_add_par(doc, "Cross-tabulation of Sociodemographic Variables by Driver Type", style = "heading 1")

# Function to add totals and format as flextable
cross_tab_with_totals <- function(var, var_label) {
  tab <- table(data$driver_type_group, var, useNA = "no")
  tab_margins <- addmargins(tab)
  
  # Rename the "Sum" row and column to "Total"
  rownames(tab_margins)[rownames(tab_margins) == "Sum"] <- "Total"
  colnames(tab_margins)[colnames(tab_margins) == "Sum"] <- "Total"
  
  # Convert to data frame and preserve rownames
  df <- as.data.frame.matrix(tab_margins)
  df <- tibble::rownames_to_column(df, var = "Driver Type")
  
  # Create flextable
  ft <- flextable(df)
  ft <- set_caption(ft, caption = paste("Cross-tabulation of", var_label, "by Driver Type"))
  ft <- autofit(ft)
  
  # Bold the "Total" row
  ft <- bold(ft, i = which(df$`Driver Type` == "Total"), bold = TRUE, part = "body")
  
  # Bold the "Total" column
  total_col_index <- which(colnames(df) == "Total")
  ft <- bold(ft, j = total_col_index, bold = TRUE, part = "body")
  
  return(ft)
}

# Loop through variables
for (var_name in names(cross_vars)) {
  var <- cross_vars[[var_name]]
  ft <- cross_tab_with_totals(var, var_name)
  doc <- body_add_par(doc, "", style = "Normal")
  doc <- body_add_flextable(doc, ft)
}

# Save document
print(doc, target = "CrossTab_With_Totals.docx")

#Graph of Demographics
library(ggplot2)
library(dplyr)

# Set custom colors for driver types
driver_colors <- c("ICE" = "#9467bd",
                   "EV"  = "#1f77b4")

# Ensure driver_type_group is a factor
data$driver_type_group <- factor(data$driver_type_group, levels = names(driver_colors))

# Get global max count across all variables
max_count <- max(
  sapply(names(cross_vars), function(var_name) {
    var_data <- cross_vars[[var_name]]
    data %>%
      mutate(var_cat = var_data) %>%
      filter(!is.na(driver_type_group) & !is.na(var_cat)) %>%
      count(driver_type_group, var_cat) %>%
      pull(n) %>%
      max()
  })
)

# Round up max_count to nearest 50 so scale looks clean
max_count <- ceiling(max_count / 50) * 50

# Loop through demographics variables and save plots
for (var_name in names(cross_vars)) {
  var_data <- cross_vars[[var_name]]
  
  plot_df <- data %>%
    select(driver_type_group) %>%
    mutate(var_cat = var_data) %>%
    filter(!is.na(driver_type_group) & !is.na(var_cat)) %>%
    group_by(driver_type_group, var_cat) %>%
    summarise(count = n(), .groups = "drop")
  
  if (tolower(var_name) == "sex") {
    # Special case for binary Male/Female as double bars
    p <- ggplot(plot_df, aes(x = var_cat, y = count, fill = driver_type_group)) +
      geom_bar(stat = "identity", position = position_dodge(width = 0.8)) +
      scale_fill_manual(values = driver_colors) +
      labs(
        title = "Distribution of Sex by Driver Type",
        x = "Sex",
        y = "Number of Respondents",
        fill = "Driver Type"
      ) +
      scale_y_continuous(limits = c(0, max_count),
                         breaks = seq(0, max_count, 50)) +
      theme_minimal() +
      theme(
        plot.title = element_text(hjust = 0.5, face = "bold", size = 16), # Title font size
        axis.title.x = element_text(size = 14),  # X-axis label size
        axis.title.y = element_text(size = 14),  # Y-axis label size
        axis.text.x  = element_text(angle = 45, hjust = 1, size = 12), # X tick labels
        axis.text.y  = element_text(size = 12)   # Y tick labels
      )
    
  } else {
    # Default for other variables
    p <- ggplot(plot_df, aes(x = var_cat, y = count, fill = driver_type_group)) +
      geom_bar(stat = "identity", position = position_dodge(width = 0.8)) +
      scale_fill_manual(values = driver_colors) +
      labs(
        title = paste("Cross tab of", var_name, "by Driver Type"),
        x = var_name,
        y = "Number of Respondents",
        fill = "Driver Type"
      ) +
      scale_y_continuous(limits = c(0, max_count),
                         breaks = seq(0, max_count, 50)) +
      theme_minimal() +
      theme(
        plot.title = element_text(hjust = 0.5, face = "bold", size = 16), # Title font size
        axis.title.x = element_text(size = 14),  # X-axis label size
        axis.title.y = element_text(size = 14),  # Y-axis label size
        axis.text.x  = element_text(angle = 45, hjust = 1, size = 12), # X tick labels
        axis.text.y  = element_text(size = 12)   # Y tick labels
      )
    
  }
  
  # Save plot
  filename <- paste0("Clustered_Bar_", var_name, ".png")
  ggsave(filename, plot = p, width = 8, height = 5, dpi = 300)
  
  cat("Saved plot:", filename, "\n")
  print(p)
}




# --- SEM with CFA, Fit Indices, Reliability & Validity Reporting ---

# Load required libraries
library(lavaan)
library(semTools)
library(semPlot)
library(dplyr)
library(flextable)
library(officer)

# --- Define SEM Model (Measurement + Structural) ---

model <- '
# Measurement model
Perceived_Usefulness =~ PU_1 + PU_2 + PU_3 + PU_4 + PU_5
Perceived_Ease_of_Use =~ PEOU_1 + PEOU_2 + PEOU_3 + PEOU_4 + PEOU_5
Willingness_to_Use =~ WTU_1 + WTU_2 + WTU_3 + WTU_4 + WTU_5
Availability_of_EV_Infrastructure =~ AOI_1 + AOI_2 + AOI_3 + AOI_4 + AOI_5
Role_of_Government_Intervention =~ ROGI_1 + ROGI_2 + ROGI_3 + ROGI_4 + ROGI_5
Environmental_Concern =~ EC_1 + EC_2 + EC_3 + EC_4 + EC_5
Perceived_Risk =~ PR_1 + PR_2 + PR_3 + PR_4 + PR_5
Social_Influence =~ SCI_1 + SCI_2 + SCI_3 + SCI_4 + SCI_5

# Structural model
Perceived_Usefulness ~ Perceived_Ease_of_Use + Availability_of_EV_Infrastructure + Role_of_Government_Intervention + Environmental_Concern + Perceived_Risk + Social_Influence
Perceived_Ease_of_Use ~ Availability_of_EV_Infrastructure + Role_of_Government_Intervention + Environmental_Concern + Perceived_Risk + Social_Influence
Willingness_to_Use ~ Perceived_Usefulness + Perceived_Ease_of_Use
'

# Subset only ICE drivers and relevant SEM variables
sem_vars <- c(
  "PU_1", "PU_2", "PU_3", "PU_4", "PU_5",
  "PEOU_1", "PEOU_2", "PEOU_3", "PEOU_4", "PEOU_5",
  "WTU_1", "WTU_2", "WTU_3", "WTU_4", "WTU_5",
  "AOI_1", "AOI_2", "AOI_3", "AOI_4", "AOI_5",
  "ROGI_1", "ROGI_2", "ROGI_3", "ROGI_4", "ROGI_5",
  "EC_1", "EC_2", "EC_3", "EC_4", "EC_5",
  "PR_1", "PR_2", "PR_3", "PR_4", "PR_5",
  "SCI_1", "SCI_2", "SCI_3", "SCI_4", "SCI_5"
)

# Subset ICE-only rows and selected SEM columns
sem_data <- thesisdata %>%
  filter(type_of_driver == "Internal Combustion Engine Vehicle driver (Patrol, diesel etc. engine vehicle)") %>%
  select(all_of(sem_vars))

sem_data <- sem_data[rowSums(is.na(sem_data)) < length(sem_vars), ]

# Fit the model with the WLSMV estimator for ordinal data
fit <- sem(model,
           data = sem_data,
           estimator = "WLSMV", # Suitable for ordinal data
           ordered = colnames(sem_data)[sapply(sem_data, is.factor)])  # Specify ordinal variables

# Check the summary
summary(fit, fit.measures = TRUE, standardized = TRUE)


# --- CFA Model Fit Export to Word ---

library(lavaan)
library(flextable)
library(officer)
library(dplyr)

# Define CFA model (measurement only)
cfa_model <- '
Perceived_Usefulness =~ PU_1 + PU_2 + PU_3 + PU_4 + PU_5
Perceived_Ease_of_Use =~ PEOU_1 + PEOU_2 + PEOU_3 + PEOU_4 + PEOU_5
Willingness_to_Use =~ WTU_1 + WTU_2 + WTU_3 + WTU_4 + WTU_5
Availability_of_EV_Infrastructure =~ AOI_1 + AOI_2 + AOI_3 + AOI_4 + AOI_5
Role_of_Government_Intervention =~ ROGI_1 + ROGI_2 + ROGI_3 + ROGI_4 + ROGI_5
Environmental_Concern =~ EC_1 + EC_2 + EC_3 + EC_4 + EC_5
Perceived_Risk =~ PR_1 + PR_2 + PR_3 + PR_4 + PR_5
Social_Influence =~ SCI_1 + SCI_2 + SCI_3 + SCI_4 + SCI_5
'

# Fit CFA model
fit_cfa <- cfa(cfa_model,
               data = sem_data,
               estimator = "WLSMV",
               ordered = colnames(sem_data)[sapply(sem_data, is.factor)])

# Extract fit indices
cfa_fit_measures <- fitMeasures(fit_cfa, c("chisq", "df", "pvalue", "cfi", "tli", "rmsea", "srmr"))

# Create a data frame for Word
cfa_fit_df <- data.frame(
  Index = c("Chi-square (χ²)", "Degrees of Freedom", "p-value", "CFI", "TLI", "RMSEA", "SRMR"),
  Value = round(unname(cfa_fit_measures), 3)
)

# Create flextable
ft_cfa <- flextable(cfa_fit_df) %>%
  set_caption("Table: CFA Model Fit Indices") %>%
  autofit() %>%
  fontsize(size = 12, part = "all") %>%
  font(fontname = "Times New Roman", part = "all") %>%
  align(align = "center", part = "all") %>%
  border_outer(part = "all", border = fp_border(color = "black")) %>%
  border_inner_h(border = fp_border(color = "black")) %>%
  border_inner_v(border = fp_border(color = "black"))

# Export to Word
doc <- read_docx() %>%
  body_add_par("CFA Model Fit Indices", style = "heading 1") %>%
  body_add_flextable(ft_cfa)

print(doc, target = "CFA_Model_Fit.docx")


library(officer)
library(flextable)

summary(fit, rsquare = TRUE)

# Extract R² values
rsq_values <- inspect(fit, "r2")
rsq_df <- data.frame(
  Variable = names(rsq_values),
  R_Squared = round(as.numeric(rsq_values), 3)
)

# Export to Word document
doc <- read_docx() %>%
  body_add_par("Table: R-squared Values for Endogenous Variables", style = "heading 1") %>%
  body_add_flextable(flextable(rsq_df))

print(doc, target = "R_squared_values.docx")

# --- Create Word Document ---
doc <- read_docx()

# --- Fit Indices Extraction ---
fit_measures <- fitMeasures(fit, c("chisq", "df", "pvalue", "cfi", "tli", "rmsea", "srmr", "aic", "bic"))
fit_df <- data.frame(
  Index = c("Chi-square (χ²)", "Degrees of Freedom", "p-value", "CFI", "TLI", "RMSEA", "SRMR", "AIC", "BIC"),
  Value = round(unname(fit_measures), 3)
)

ft_fit <- flextable(fit_df) %>%
  set_caption("Table 1: Model Fit Indices") %>%
  autofit() %>%
  fontsize(size = 12, part = "all") %>%
  font(fontname = "Times New Roman", part = "all") %>%
  flextable::align(align = "center", part = "all") %>%
  border_remove() %>%
  border_outer(part = "all", border = fp_border(color = "black")) %>%
  border_inner_h(border = fp_border(color = "black")) %>%
  border_inner_v(border = fp_border(color = "black"))

doc <- body_add_par(doc, "Model Fit Indices", style = "heading 1")
doc <- body_add_flextable(doc, ft_fit)
doc <- body_add_par(doc, "", style = "Normal")
#---------------------------------------------------

# ---To Compute Cronbach's Alpha for Each Latent Variable ---

library(psych)

# Extract latent constructs
latent_vars <- lavNames(fit, type = "lv")

# Get the dataset used in the SEM
data_used <- lavInspect(fit, "data")

# Get measurement model
model_params <- parameterEstimates(fit) %>% filter(op == "=~")

# Prepare list to store alphas
alpha_list <- list()

# Loop over latent constructs to compute alpha
for (lv in latent_vars) {
  indicators <- model_params %>%
    filter(lhs == lv) %>%
    pull(rhs)
  
  data_subset <- data_used[, indicators, drop = FALSE]
  
  # Compute Cronbach's alpha
  alpha_result <- psych::alpha(data_subset)
  alpha_value <- round(alpha_result$total$raw_alpha, 3)
  
  alpha_list[[lv]] <- alpha_value
}

# Convert to data frame
alpha_df <- data.frame(
  Construct = names(alpha_list),
  Cronbach_Alpha = unlist(alpha_list),
  row.names = NULL
)

# Create flextable
ft_alpha <- flextable(alpha_df) %>%
  set_caption("Table X: Cronbach's Alpha for Internal Consistency Reliability") %>%
  autofit() %>%
  fontsize(size = 12, part = "all") %>%
  font(fontname = "Times New Roman", part = "all") %>%
  flextable::align(align = "center", part = "all") %>%
  border_remove() %>%
  border_outer(border = fp_border(color = "black")) %>%
  border_inner_h(border = fp_border(color = "black")) %>%
  border_inner_v(border = fp_border(color = "black"))

# Add to Word document
doc <- body_add_par(doc, "Internal Consistency Reliability – Cronbach's Alpha", style = "heading 1")
doc <- body_add_flextable(doc, ft_alpha)
doc <- body_add_par(doc, "", style = "Normal")

library(psych)
library(dplyr)
library(flextable)
library(lavaan)

# Extract latent constructs
latent_vars <- lavNames(fit, type = "lv")

# Get the dataset used in the SEM
data_used <- lavInspect(fit, "data")

# Get measurement model
model_params <- parameterEstimates(fit) %>% filter(op == "=~")

# Prepare list to store item-level alphas
item_alpha_list <- list()

# Loop over latent constructs to compute item-level alpha
for (lv in latent_vars) {
  indicators <- model_params %>%
    filter(lhs == lv) %>%
    pull(rhs)
  
  # Loop through each indicator (item)
  for (indicator in indicators) {
    # Exclude the current indicator to compute Cronbach's alpha for remaining items
    remaining_items <- setdiff(indicators, indicator)
    
    data_subset <- data_used[, remaining_items, drop = FALSE]
    
    # Compute Cronbach's alpha for the remaining items
    alpha_result <- psych::alpha(data_subset)
    alpha_value <- round(alpha_result$total$raw_alpha, 3)
    
    # Store the result with item and corresponding alpha
    item_alpha_list[[length(item_alpha_list) + 1]] <- data.frame(
      Item = indicator,
      Cronbach_Alpha = alpha_value
    )
  }
}

# Combine all results into a single data frame
item_alpha_df <- do.call(rbind, item_alpha_list)

# Create flextable for item-level alphas
ft_item_alpha <- flextable(item_alpha_df) %>%
  set_caption("Table X: Item-Level Cronbach's Alpha for Internal Consistency Reliability") %>%
  autofit() %>%
  fontsize(size = 12, part = "all") %>%
  font(fontname = "Times New Roman", part = "all") %>%
  flextable::align(align = "center", part = "all") %>%  # Explicitly call align from flextable
  border_remove() %>%
  border_outer(border = fp_border(color = "black")) %>%
  border_inner_h(border = fp_border(color = "black")) %>%
  border_inner_v(border = fp_border(color = "black"))


# Add to Word document
doc <- body_add_par(doc, "Internal Consistency Reliability – Item-Level Cronbach's Alpha", style = "heading 1")
doc <- body_add_flextable(doc, ft_item_alpha)
doc <- body_add_par(doc, "", style = "Normal")


# --- CFA Loadings ---
cfa_loadings <- parameterEstimates(fit, standardized = TRUE) %>%
  filter(op == "=~") %>%
  select(Latent = lhs, Indicator = rhs, Std_Loading = std.all)

ft_cfa <- flextable(cfa_loadings) %>%
  set_caption("Table 2: Confirmatory Factor Analysis – Standardized Loadings") %>%
  autofit() %>%
  fontsize(size = 12, part = "all") %>%
  font(fontname = "Times New Roman", part = "all") %>%
  flextable::align(align = "center", part = "all") %>%  # Explicit call
  border_remove() %>%
  border_outer(part = "all", border = fp_border(color = "black")) %>%
  border_inner_h(border = fp_border(color = "black")) %>%
  border_inner_v(border = fp_border(color = "black"))


doc <- body_add_par(doc, "Confirmatory Factor Analysis Results", style = "heading 1")
doc <- body_add_flextable(doc, ft_cfa)
doc <- body_add_par(doc, "", style = "Normal")

# --- Reliability & Convergent Validity ---
rel_stats <- compRelSEM(fit)
ave_stats <- AVE(fit)
constructs <- intersect(names(rel_stats), names(ave_stats))

rel_valid_df <- data.frame(
  Construct = constructs,
  Composite_Reliability = round(unlist(rel_stats[constructs]), 3),
  AVE = round(unlist(ave_stats[constructs]), 3),
  row.names = NULL
)

ft_rel <- flextable(rel_valid_df) %>%
  set_caption("Table 3: Reliability and Convergent Validity (CR & AVE)") %>%
  autofit() %>%
  fontsize(size = 12, part = "all") %>%
  font(fontname = "Times New Roman", part = "all") %>%
  flextable::align(align = "center", part = "all") %>%  # Use explicit namespace
  border_remove() %>%
  border_outer(border = fp_border(color = "black")) %>%  # Removed part = "all"
  border_inner_h(border = fp_border(color = "black")) %>%
  border_inner_v(border = fp_border(color = "black"))

doc <- body_add_par(doc, "Reliability and Convergent Validity Results", style = "heading 1")
doc <- body_add_flextable(doc, ft_rel)
doc <- body_add_par(doc, "", style = "Normal")

# --- Discriminant Validity: Fornell-Larcker Criterion (All Constructs) ---
all_lvs <- lavNames(fit, type = "lv")
ave_vals_all <- AVE(fit)
ave_sqrt_all <- sqrt(ave_vals_all[all_lvs])
cor_lv_all <- inspect(fit, "cor.lv")
diag(cor_lv_all) <- ave_sqrt_all
fornell_all_df <- round(as.data.frame(cor_lv_all), 3)
fornell_all_df <- cbind(Construct = rownames(fornell_all_df), fornell_all_df)

ft_fornell_all <- flextable(fornell_all_df) %>%
  set_caption("Table 4: Discriminant Validity – Fornell-Larcker Criterion (All Constructs)") %>%
  autofit() %>%
  fontsize(size = 12, part = "all") %>%
  font(fontname = "Times New Roman", part = "all") %>%
  flextable::align(align = "center", part = "all") %>%  # Explicit namespacing
  border_remove() %>%
  border_outer(border = fp_border(color = "black")) %>%  # Removed invalid `part`
  border_inner_h(border = fp_border(color = "black")) %>%
  border_inner_v(border = fp_border(color = "black"))

for (i in seq_along(all_lvs)) {
  lv <- all_lvs[i]
  ft_fornell_all <- bold(ft_fornell_all, j = lv, i = i, bold = TRUE)
}

doc <- body_add_par(doc, "Discriminant Validity for All Constructs", style = "heading 1")
doc <- body_add_flextable(doc, ft_fornell_all)
doc <- body_add_par(doc, "", style = "Normal")

# --- Discriminant Validity: Endogenous Constructs Only ---
paths <- parameterEstimates(fit)
endo_lvs <- unique(paths$lhs[paths$op == "~" & paths$lhs %in% all_lvs])
ave_vals <- AVE(fit)
ave_sqrt <- sqrt(ave_vals[endo_lvs])
cor_matrix <- inspect(fit, "cor.lv")[endo_lvs, endo_lvs]
diag(cor_matrix) <- ave_sqrt
fornell_df <- round(as.data.frame(cor_matrix), 3)
fornell_df <- cbind(Construct = rownames(fornell_df), fornell_df)

ft_fornell <- flextable(fornell_df) %>%
  set_caption("Table 5: Discriminant Validity – Fornell-Larcker (Endogenous Constructs Only)") %>%
  autofit() %>%
  fontsize(size = 12, part = "all") %>%
  font(fontname = "Times New Roman", part = "all") %>%
  flextable::align(align = "center", part = "all") %>%
  border_remove() %>%
  border_outer(border = fp_border(color = "black")) %>%  # Removed `part = "all"`
  border_inner_h(border = fp_border(color = "black")) %>%
  border_inner_v(border = fp_border(color = "black"))

for (i in seq_along(endo_lvs)) {
  ft_fornell <- bold(ft_fornell, j = endo_lvs[i], i = i, bold = TRUE)
}

doc <- body_add_par(doc, "Discriminant Validity for Endogenous Constructs", style = "heading 1")
doc <- body_add_flextable(doc, ft_fornell)
doc <- body_add_par(doc, "", style = "Normal")

# --- Discriminant Validity: Exogenous Constructs Only ---
exo_lvs <- setdiff(all_lvs, endo_lvs)
ave_vals_exo <- AVE(fit)
ave_sqrt_exo <- sqrt(ave_vals_exo[exo_lvs])
cor_exo <- inspect(fit, "cor.lv")[exo_lvs, exo_lvs]
diag(cor_exo) <- ave_sqrt_exo
fornell_exo_df <- round(as.data.frame(cor_exo), 3)
fornell_exo_df <- cbind(Construct = rownames(fornell_exo_df), fornell_exo_df)

ft_fornell_exo <- flextable(fornell_exo_df) %>%
  set_caption("Table 6: Discriminant Validity – Fornell-Larcker (Exogenous Constructs Only)") %>%
  autofit() %>%
  fontsize(size = 12, part = "all") %>%
  font(fontname = "Times New Roman", part = "all") %>%
  flextable::align(align = "center", part = "all") %>%  # Namespaced align()
  border_remove() %>%
  border_outer(border = fp_border(color = "black")) %>%  # Removed invalid `part`
  border_inner_h(border = fp_border(color = "black")) %>%
  border_inner_v(border = fp_border(color = "black"))

for (i in seq_along(exo_lvs)) {
  ft_fornell_exo <- bold(ft_fornell_exo, j = exo_lvs[i], i = i, bold = TRUE)
}

doc <- body_add_par(doc, "Discriminant Validity for Exogenous Constructs", style = "heading 1")
doc <- body_add_flextable(doc, ft_fornell_exo)
doc <- body_add_par(doc, "", style = "Normal")

# --- Structural Path Coefficients (Endo and Exo) ---
std_est <- standardizedSolution(fit)
structural_paths <- std_est %>%
  filter(op == "~") %>%
  select(Predictor = rhs, Outcome = lhs, Beta = est.std) %>%
  mutate(Beta = round(Beta, 3))

endogenous_vars <- c("Perceived_Usefulness", "Perceived_Ease_of_Use", "Willingness_to_Use")
structural_paths <- structural_paths %>%
  mutate(Type = ifelse(Predictor %in% endogenous_vars, "Endogenous to Endogenous", "Exogenous to Endogenous"))

endo_table <- structural_paths %>% filter(Type == "Endogenous to Endogenous") %>% select(-Type)
exo_table <- structural_paths %>% filter(Type == "Exogenous to Endogenous") %>% select(-Type)

ft_endo <- flextable(endo_table) %>%
  set_caption("Table 7a: Endogenous-to-Endogenous Path Coefficients") %>%
  autofit() %>%
  fontsize(size = 12, part = "all") %>%
  font(fontname = "Times New Roman", part = "all") %>%
  flextable::align(align = "center", part = "all") %>%
  border_remove() %>%
  border_outer(border = fp_border(color = "black")) %>%
  border_inner_h(border = fp_border(color = "black")) %>%
  border_inner_v(border = fp_border(color = "black"))

ft_exo <- flextable(exo_table) %>%
  set_caption("Table 7b: Exogenous-to-Endogenous Path Coefficients") %>%
  autofit() %>%
  fontsize(size = 12, part = "all") %>%
  font(fontname = "Times New Roman", part = "all") %>%
  flextable::align(align = "center", part = "all") %>%  # Use namespaced align()
  border_remove() %>%
  border_outer(border = fp_border(color = "black")) %>%
  border_inner_h(border = fp_border(color = "black")) %>%
  border_inner_v(border = fp_border(color = "black"))

doc <- body_add_par(doc, "Standardized Path Coefficients", style = "heading 1")
doc <- body_add_flextable(doc, ft_endo)
doc <- body_add_par(doc, "")
doc <- body_add_flextable(doc, ft_exo)

#---------------------------

# Load required libraries
library(dplyr)
library(tidyr)
library(flextable)

# Extract the original dataset used in fitting the model
data_used <- lavInspect(fit, "data")  # raw data

# Extract the indicators for each latent variable
mod_summary <- parameterEstimates(fit, standardized = TRUE)
loadings <- mod_summary %>%
  filter(op == "=~") %>%
  select(latent = lhs, observed = rhs, loading = std.all)

# Get unique observed variables used in the measurement model
observed_vars <- unique(loadings$observed)

# Compute observed correlations
cor_obs <- cor(data_used[, observed_vars], use = "pairwise.complete.obs")

# HTMT computation
latent_vars <- unique(loadings$latent)
htmt_mat <- matrix(NA, nrow = length(latent_vars), ncol = length(latent_vars),
                   dimnames = list(latent_vars, latent_vars))

for (i in 1:length(latent_vars)) {
  for (j in 1:length(latent_vars)) {
    if (i < j) {
      lv1 <- latent_vars[i]
      lv2 <- latent_vars[j]
      
      items1 <- loadings %>% filter(latent == lv1) %>% pull(observed)
      items2 <- loadings %>% filter(latent == lv2) %>% pull(observed)
      
      inter_trait_corrs <- cor_obs[items1, items2]
      htmt_val <- mean(abs(inter_trait_corrs), na.rm = TRUE)
      htmt_mat[i, j] <- htmt_val
      htmt_mat[j, i] <- htmt_val
    }
  }
}
diag(htmt_mat) <- 1  # set diagonals to 1

# Format HTMT as a dataframe
htmt_df <- as.data.frame(round(htmt_mat, 3)) %>%
  mutate(Construct = rownames(.)) %>%
  select(Construct, everything())

# Create flextable
ft_htmt <- flextable(htmt_df) %>%
  set_caption("Table X: Discriminant Validity – HTMT Matrix") %>%
  autofit() %>%
  fontsize(size = 12, part = "all") %>%
  font(fontname = "Times New Roman", part = "all") %>%
  flextable::align(align = "center", part = "all") %>%  # Use explicit namespace
  border_remove() %>%
  border_outer(border = fp_border(color = "black")) %>%  # Removed `part = "all"`
  border_inner_h(border = fp_border(color = "black")) %>%
  border_inner_v(border = fp_border(color = "black"))


# Add to Word document
doc <- body_add_par(doc, "Discriminant Validity – HTMT Matrix (Manually Computed)", style = "heading 1")
doc <- body_add_flextable(doc, ft_htmt)
doc <- body_add_par(doc, "", style = "Normal")

#Descriptive Statistics of Likert scale of ICE

# Load necessary libraries
library(psych)       # for descriptive statistics
library(dplyr)       # for data manipulation
library(flextable)   # for Word export
library(officer)     # for Word document handling

# Step 1: Extract original data from fitted SEM model
data_used <- lavInspect(fit, "data")

# Step 2: Get observed (indicator) variables from the measurement model
loadings <- parameterEstimates(fit) %>%
  filter(op == "=~") %>%
  select(latent = lhs, observed = rhs)

# Step 3: Identify the unique Likert scale items (observed indicators)
likert_items <- unique(loadings$observed)

# Step 4: Subset dataset to only include Likert items
likert_data <- data_used[, likert_items]

# Step 5: Compute descriptive statistics
desc_stats <- psych::describe(likert_data)[, c("n", "mean", "sd", "skew", "kurtosis")]
desc_stats <- round(desc_stats, 2)
desc_stats <- as.data.frame(desc_stats)
desc_stats <- tibble::rownames_to_column(desc_stats, "Item")

# Step 6: Format output as flextable
ft_desc <- flextable(desc_stats) %>%
  set_caption("Table X: Descriptive Statistics of Likert Scale Items") %>%
  autofit() %>%
  fontsize(size = 12, part = "all") %>%
  font(fontname = "Times New Roman", part = "all") %>%
  flextable::align(align = "center", part = "all") %>%  # Namespaced align()
  border_remove() %>%
  border_outer(border = fp_border(color = "black")) %>%  # Removed invalid part
  border_inner_h(border = fp_border(color = "black")) %>%
  border_inner_v(border = fp_border(color = "black"))

# Step 7: Add to Word document
doc <- body_add_par(doc, "Descriptive Statistics for Likert Scale Items", style = "heading 1")
doc <- body_add_flextable(doc, ft_desc)
doc <- body_add_par(doc, "", style = "Normal")

# Load necessary libraries
library(lavaan)
library(dplyr)
library(flextable)
library(officer)

# --- Extract Standardized Estimates, SEs, p-values, and 95% CIs ---
std_table <- standardizedSolution(fit) %>%
  filter(op == "~") %>%  # Only structural paths (regressions)
  select(lhs, rhs, est.std, se, pvalue, ci.lower, ci.upper) %>%
  rename(
    Outcome = lhs,
    Predictor = rhs,
    `Standardized Estimate (β)` = est.std,
    `Standard Error (SE)` = se,
    `p-value` = pvalue,
    `95% CI Lower` = ci.lower,
    `95% CI Upper` = ci.upper
  ) %>%
  mutate(across(`Standardized Estimate (β)`:`95% CI Upper`, ~round(., 3)))  # Round numeric columns

# --- Format the Table ---
ft_paths <- flextable(std_table) %>%
  set_caption("Table 4: Structural Path Coefficients – Standardized Estimates, SEs, CIs, and Significance") %>%
  autofit() %>%
  fontsize(size = 12, part = "all") %>%
  font(fontname = "Times New Roman", part = "all") %>%
  flextable::align(align = "center", part = "all") %>%
  border_remove() %>%
  border_outer(border = fp_border(color = "black")) %>%
  border_inner_h(border = fp_border(color = "black")) %>%
  border_inner_v(border = fp_border(color = "black"))

# --- Add heading and table to the Word document ---
doc <- body_add_par(doc, "Structural Model Results", style = "heading 1")
doc <- body_add_flextable(doc, ft_paths)
doc <- body_add_par(doc, "", style = "Normal")  # Line space


# --- Save Word Document ---
print(doc, target = "ICE_SEM_Analysis_Report.docx")



#Visualization of Likert Scale Plots

#================================================================================
#Likert Plot for ICE drivers

library(ggplot2)
library(forcats)
library(dplyr)
library(tidyr)
library(purrr)
library(stringr)
library(RColorBrewer)
library(likert)

# 1. CONSTRUCT DEFINITIONS ----------------------------------------------------
constructs <- list(
  "Perceived Usefulness"              = c("PU_1","PU_2","PU_3","PU_4","PU_5"),
  "Perceived Ease of Use"             = c("PEOU_1","PEOU_2","PEOU_3","PEOU_4","PEOU_5"),
  "Willingness to Use"                = c("WTU_1","WTU_2","WTU_3","WTU_4","WTU_5"),
  "Availability of EV Infrastructure" = c("AOI_1","AOI_2","AOI_3","AOI_4","AOI_5"),
  "Role of Government Intervention"   = c("ROGI_1","ROGI_2","ROGI_3","ROGI_4","ROGI_5"),
  "Environmental Concern"             = c("EC_1","EC_2","EC_3","EC_4","EC_5"),
  "Perceived Risk"                    = c("PR_1","PR_2","PR_3","PR_4","PR_5"),
  "Socio-Cultural Influence"          = c("SCI_1","SCI_2","SCI_3","SCI_4","SCI_5")
)

# 2. OUTPUT PATH --------------------------------------------------------------
out_dir <- "C:/Users/Dominic Ocloo/Documents/Thesis_Analysis"
dir.create(out_dir, showWarnings = FALSE, recursive = TRUE)

# 3. SHARED SETTINGS ----------------------------------------------------------
lk_levels <- c("Strongly Disagree", "Disagree",
               "Neutral", "Agree", "Strongly Agree")
pal <- RColorBrewer::brewer.pal(5, "RdYlGn")
names(pal) <- lk_levels                          # keep mapping explicit

# 4. COERCION HELPER ----------------------------------------------------------
coerce_likert <- function(x) {
  if (is.factor(x)) {
    # make sure it carries the full level set
    x <- forcats::fct_expand(x, lk_levels)
    x <- factor(x, levels = lk_levels, ordered = TRUE)
  } else {
    # numeric or character codes
    x <- suppressWarnings(as.integer(x))
    x <- factor(x, levels = 1:5, labels = lk_levels, ordered = TRUE)
  }
  levels(x) <- lk_levels
  x
}

# 5. LEGEND LABEL BUILDER -----------------------------------------------------
legend_labels <- function(df) {
  df_long <- tidyr::pivot_longer(df, everything(),
                                 names_to = "Item", values_to = "Resp") |>
    filter(!is.na(Resp)) |>
    count(Resp, name = "n") |>
    mutate(pct = round(100 * n / sum(n), 1),
           lbl = paste0(Resp, " (", pct, "%)"))
  setNames(df_long$lbl, df_long$Resp)
}

# 6. PLOT ONE CONSTRUCT -------------------------------------------------------
plot_construct <- function(data, vars, title_txt) {
  df <- data |>
    select(all_of(vars)) |>
    mutate(across(everything(), coerce_likert)) |>
    as.data.frame()
  
  lab_vec <- legend_labels(df)
  lik_obj <- likert::likert(df)
  
  p <- plot(lik_obj, facet = FALSE,
            low.color = pal[1], high.color = pal[5]) +
    scale_fill_manual(values = pal,
                      breaks = lk_levels,
                      labels = lab_vec,
                      name   = "Response") +
    labs(title = title_txt) +
    theme_minimal(base_family = "sans") +
    theme(axis.title = element_blank())
  
  print(p)
  ggsave(file.path(out_dir,
                   paste0("Likert_", str_replace_all(title_txt, "\\s+", "_"),
                          ".png")),
         p, width = 10, height = 6, dpi = 300)
}

# 7. INDIVIDUAL CONSTRUCT GRAPHS ---------------------------------------------
for (ct in names(constructs)) {
  vars <- constructs[[ct]]
  if (all(vars %in% names(thesisdata))) {
    plot_construct(thesisdata, vars, ct)
  } else {
    warning("Missing items in construct: ", ct)
  }
}

## 8. COMBINED GRAPH (NAs EXCLUDED) -------------------------------------------
combined_long <- purrr::imap_dfr(
  constructs,
  \(vars, cname) {
    thesisdata |>
      select(all_of(vars)) |>
      mutate(across(everything(), coerce_likert)) |>
      mutate(row_id = row_number()) |>
      pivot_longer(-row_id, names_to = "Item", values_to = "Resp") |>
      mutate(Construct = cname)
  }
) |>
  filter(!is.na(Resp))

# summarise %
summary_df <- combined_long |>
  count(Construct, Resp) |>
  group_by(Construct) |>
  mutate(Percent = 100 * n / sum(n)) |>
  ungroup()

# average legend %
legend_df <- summary_df |>
  group_by(Resp) |>
  summarise(avg = round(mean(Percent), 1),
            .groups = "drop") |>
  mutate(lbl = paste0(Resp, " (", avg, "%)"))
comb_labels <- setNames(legend_df$lbl, legend_df$Resp)

summary_df$Construct <- factor(summary_df$Construct, levels = names(constructs))

p_comb <- ggplot(summary_df,
                 aes(x = Construct, y = Percent, fill = Resp)) +
  geom_bar(stat = "identity", width = .8) +
  coord_flip() +
  scale_fill_manual(values = pal,
                    breaks = lk_levels,
                    labels = comb_labels,
                    name   = "Response") +
  labs(title = "Overall Likert Distribution by Construct",
       y = "Percentage", x = NULL) +
  theme_minimal(base_family = "sans") +
  theme(legend.position = "right")

print(p_comb)
ggsave(file.path(out_dir, "Likert_All_Constructs_Combined.png"),
       p_comb, width = 11, height = 7, dpi = 300)

#Check folder holding thesis analysis




#-----------------------------------------------------------------

# VISUALIZE MODEL IN SEMPLOT
# Load required libraries
library(lavaan)
library(semPlot)

# Get original latent variable names
latents <- lavNames(fit, type = "lv")

# Generate initials for latent variables (e.g., PU, PEOU, etc.)
initials <- sapply(strsplit(latents, "_"), function(x) paste0(substr(x, 1, 1), collapse = ""))
label_map <- setNames(initials, latents)

# Extract semPlot model object
sem_model <- semPlot::semPlotModel(fit)

# Extract original node names
all_labels <- sem_model@Vars$name

# Replace only latent variable names with initials
new_labels <- ifelse(all_labels %in% names(label_map),
                     label_map[all_labels],
                     all_labels)

# Set custom colors for latent variables
latent_colors <- rep("#1f77b4", length(latents))
names(latent_colors) <- latents
latent_colors["Willingness_to_Use"] <- "#2ca02c"
latent_colors["Perceived_Usefulness"] <- "#ff7f0e"
latent_colors["Perceived_Ease_of_Use"] <- "#ff7f0e"

# Draw the path diagram with custom labels
semPaths(
  object = fit,
  what = "std",
  layout = "tree",
  style = "ram",
  rotation = 1,
  sizeMan = 4,
  sizeLat = 8,
  edge.label = NULL,
  exoVar = TRUE,
  intercepts = FALSE,
  residuals = FALSE,
  color = list(lat = latent_colors, man = "#9467bd"),
  edge.color = "black",
  edge.width = 0.5,
  optimizeLatRes = TRUE,
  nCharNodes = 0,
  title = FALSE,
  mar = c(2, 2, 2, 2),
  nodeLabels = new_labels  # Use modified labels (initials)
)
#---------------------------------------------------------------------------
#Model excluding item loadings

# Load required libraries
library(lavaan)
library(semPlot)

# Extract latent variable names and create initials
latents <- lavNames(fit, type = "lv")
initials <- sapply(strsplit(latents, "_"), function(x) paste(substr(x, 1, 1), collapse = ""))
names(initials) <- latents

# Assign colors: green for 'Willingness_to_Use', orange for PU and PEOU
latent_colors <- rep("#1f77b4", length(latents))
if ("Willingness_to_Use" %in% latents) {
  latent_colors[latents == "Willingness_to_Use"] <- "#2ca02c"
}
if ("Perceived_Usefulness" %in% latents) {
  latent_colors[latents == "Perceived_Usefulness"] <- "#ff7f0e"
}
if ("Perceived_Ease_of_Use" %in% latents) {
  latent_colors[latents == "Perceived_Ease_of_Use"] <- "#ff7f0e"
}

# Plot latent-only structural model with smooth, curved arrows
semPaths(
  object = fit,
  what = "std",
  whatLabels = "std",
  style = "ram",
  layout = "tree2",               # Alternative tree layout with better spacing
  rotation = 1,                   # Top-down
  sizeLat = 8,
  sizeMan = 0,                    # Hide observed variables
  nodeLabels = initials,          # Use initials
  nCharNodes = 0,                 # No truncation
  color = list(lat = latent_colors),
  edge.color = "black",
  edge.width = 0.7,
  curve = 2.0,                    # Increase curvature for less overlap
  structural = TRUE,              # Only show structural (latent-to-latent) paths
  residuals = FALSE,
  intercepts = FALSE,
  optimizeLatRes = TRUE,
  mar = c(3, 7, 7, 7),
  title = FALSE
)

#===============================================================================

#EV drivers model

# Load required packages
library(lavaan)
library(dplyr)
library(psych)
library(flextable)
library(officer)
library(semPlot)

# --- 1. Fit Endogenous SEM Model ---
model_endo <- '
  Perceived_Usefulness    =~ PU_A + PU_B + PU_C + PU_D + PU_E
  Perceived_Ease_of_Use    =~ PEOU_A + PEOU_B + PEOU_C + PEOU_D + PEOU_E
  Willingness_to_Use       =~ WTU_A + WTU_B + WTU_C + WTU_D + WTU_E
  Actual_Usage             =~ AU_A  + AU_B  + AU_C  + AU_D  + AU_E

  Willingness_to_Use ~ Perceived_Usefulness + Perceived_Ease_of_Use
  Actual_Usage       ~ Willingness_to_Use
'

fit_endo <- sem(
  model    = model_endo,
  data     = thesisdata,
  estimator= "WLSMV",
  ordered  = colnames(thesisdata)[sapply(thesisdata, is.factor)]
)

# Visualize the endogenous structural model

# 1. Define mapping from full names to initials
label_map <- c(
  "Perceived_Usefulness"     = "PU",
  "Perceived_Ease_of_Use"    = "PEOU",
  "Willingness_to_Use"       = "WTU",
  "Actual_Usage"             = "AU"
)

# 2. Extract model structure
sem_model <- semPlot::semPlotModel(fit_endo)

# 3. Shorten latent variable names using the map
short_node_labels <- sapply(sem_model@Vars$name, function(x) {
  if (x %in% names(label_map)) {
    label_map[x]
  } else {
    x  # keep indicators as-is
  }
})

# 4. Assign node colors: green for AU, orange for others, yellow for indicators
latent_vars <- lavNames(fit_endo, type = "lv")
indicator_vars <- lavNames(fit_endo, type = "ov")

latent_colors <- rep("darkorange", length(latent_vars))
names(latent_colors) <- latent_vars
latent_colors["Actual_Usage"] <- "forestgreen"

indicator_colors <- rep("#9467bd", length(indicator_vars))
names(indicator_colors) <- indicator_vars

all_nodes <- sem_model@Vars$name
node_color_vec <- ifelse(all_nodes %in% names(latent_colors),
                         latent_colors[all_nodes],
                         indicator_colors[all_nodes])

# 5. Plot with arrows in black and updated labels
semPaths(
  fit_endo,
  what = "std",
  layout = "tree",
  edge.label.cex = 0.9,
  sizeMan = 6,
  sizeLat = 8,
  edge.color = "black",         # All arrows black
  fade = FALSE,
  style = "lisrel",
  residuals = FALSE,
  intercepts = FALSE,
  color = node_color_vec,
  nodeLabels = short_node_labels,  # Initials for constructs
  borders = FALSE
)

#-------------------------------------------------------------------------------
#Model without Standard loadings
  
library(semPlot)
library(lavaan)

# ---- 1. Define latent variable labels ----
label_map <- c(
  "Perceived_Usefulness"     = "PU",
  "Perceived_Ease_of_Use"    = "PEOU",
  "Willingness_to_Use"       = "WTU",
  "Actual_Usage"             = "AU"
)

# ---- 2. Extract model structure ----
sem_model <- semPlot::semPlotModel(fit_endo)
latent_vars   <- lavNames(fit_endo, type = "lv")
indicator_vars <- lavNames(fit_endo, type = "ov")
all_nodes <- sem_model@Vars$name

# ---- 3. Set short labels for latent variables ----
short_node_labels <- sapply(all_nodes, function(x) {
  if (x %in% names(label_map)) label_map[[x]] else ""
})

# ---- 4. Define node colors ----
node_color_vec <- rep("transparent", length(all_nodes))
names(node_color_vec) <- all_nodes
node_color_vec[latent_vars] <- "darkorange"
node_color_vec["Actual_Usage"] <- "forestgreen"

# ---- 5. Plot model (constructs only with beta values) ----
semPaths(
  fit_endo,
  what        = "paths",         # structural model only
  whatLabels  = "std",           # show standardized beta values
  layout      = "tree",
  style       = "lisrel",
  residuals   = FALSE,
  intercepts  = FALSE,
  fade        = FALSE,
  sizeMan     = 0,               # hide observed indicators
  sizeLat     = 8,
  edge.color  = "black",
  edge.label.cex = 1.1,          # font size for beta values
  color       = node_color_vec,
  nodeLabels  = short_node_labels,
  mar         = c(5, 5, 5, 5)
)

  
  
#--------------------------

doc <- read_docx()


# ===============================================================
# 1. CFA-ONLY MODEL FIT FOR ENDOGENOUS MODEL
# ===============================================================

model_endo_cfa <- '
  Perceived_Usefulness    =~ PU_A + PU_B + PU_C + PU_D + PU_E
  Perceived_Ease_of_Use   =~ PEOU_A + PEOU_B + PEOU_C + PEOU_D + PEOU_E
  Willingness_to_Use      =~ WTU_A + WTU_B + WTU_C + WTU_D + WTU_E
  Actual_Usage            =~ AU_A  + AU_B  + AU_C  + AU_D  + AU_E
'

fit_endo_cfa <- cfa(
  model = model_endo_cfa,
  data = thesisdata,
  estimator = "WLSMV",
  ordered = colnames(thesisdata)[sapply(thesisdata, is.factor)]
)

fit_indices_endo_cfa <- fitMeasures(
  fit_endo_cfa,
  c("chisq","df","pvalue","cfi","tli","rmsea","srmr")
)

fit_endo_cfa_df <- data.frame(
  Index = c("Chi-square (χ²)", "df", "p-value", "CFI", "TLI", "RMSEA", "SRMR"),
  Value = round(unname(fit_indices_endo_cfa), 3)
)

ft_endo_cfa <- flextable(fit_endo_cfa_df) %>%
  set_caption("Table: CFA Model Fit Indices (Endogenous Model)") %>%
  autofit()

doc <- doc %>%
  body_add_par("CFA Model Fit (Endogenous Measurement Model)", style="heading 1") %>%
  body_add_flextable(ft_endo_cfa) %>%
  body_add_par("", style="Normal")


# ===============================================================
# 2. SEM MODEL FIT
# ===============================================================

fit_indices_endo <- fitMeasures(
  fit_endo,
  c("chisq","df","pvalue","cfi","tli","rmsea","srmr")
)

fit_endo_df <- data.frame(
  Index = c("Chi-square (χ²)", "df", "p-value", "CFI", "TLI", "RMSEA", "SRMR"),
  Value = round(unname(fit_indices_endo), 3)
)

ft_fit_endo <- flextable(fit_endo_df) %>%
  set_caption("Table: SEM Model Fit Indices (Endogenous Model)") %>%
  autofit()

doc <- doc %>%
  body_add_par("SEM Model Fit", style="heading 1") %>%
  body_add_flextable(ft_fit_endo) %>%
  body_add_par("", style="Normal")


# ===============================================================
# 3. CFA LOADINGS (Correct: use fit_endo_cfa)
# ===============================================================

loadings_df <- parameterEstimates(fit_endo_cfa, standardized=TRUE) %>%
  filter(op == "=~") %>%
  select(Latent = lhs, Indicator = rhs, Std_Loading = std.all) %>%
  mutate(Std_Loading = round(Std_Loading, 3))

ft_cfa_loadings <- flextable(loadings_df) %>%
  set_caption("Table: CFA Standardized Loadings (Endogenous Model)") %>%
  autofit()

doc <- doc %>%
  body_add_par("Standardized CFA Loadings", style="heading 1") %>%
  body_add_flextable(ft_cfa_loadings) %>%
  body_add_par("", style="Normal")


# ===============================================================
# 4. CRONBACH’S ALPHA (Construct Level) — FIXED VERSION
# ===============================================================

latent_vars <- lavNames(fit_endo_cfa, type="lv")
model_params <- parameterEstimates(fit_endo_cfa) %>% filter(op == "=~")
data_used <- lavInspect(fit_endo_cfa, "data")

alpha_results <- list()

for (lv in latent_vars) {
  
  # Indicators for this latent variable
  inds <- model_params %>% filter(lhs == lv) %>% pull(rhs)
  ds <- data_used[, inds, drop = FALSE]
  
  # Convert to numeric safely
  ds[] <- lapply(ds, function(x) suppressWarnings(as.numeric(as.character(x))))
  
  # Check if conversion produced numeric data
  if (!all(sapply(ds, is.numeric))) {
    alpha_results[[lv]] <- NA
    next
  }
  
  # Check for missing or constant values
  if (any(sapply(ds, function(x) length(unique(x)) <= 1))) {
    alpha_results[[lv]] <- NA
    next
  }
  
  # Compute Cronbach's alpha safely
  a <- try(psych::alpha(ds)$total$raw_alpha, silent = TRUE)
  
  if (inherits(a, "try-error")) {
    alpha_results[[lv]] <- NA
  } else {
    alpha_results[[lv]] <- a
  }
}

# Create final dataframe
alpha_df <- data.frame(
  Construct = names(alpha_results),
  Cronbach_Alpha = round(unlist(alpha_results), 3)
)

# Create flextable
ft_alpha <- flextable(alpha_df) %>%
  set_caption("Table: Construct-Level Cronbach's Alpha (Corrected)") %>%
  autofit()


# ===============================================================
# 5. COMPOSITE RELIABILITY & AVE
# ===============================================================

std_est <- standardizedSolution(fit_endo_cfa) %>% filter(op == "=~")
constructs <- unique(std_est$lhs)

cr <- ave <- numeric(length(constructs))
names(cr) <- names(ave) <- constructs

for (c in constructs) {
  loads <- std_est %>% filter(lhs == c) %>% pull(est.std)
  errs <- 1 - loads^2
  cr[c]  <- (sum(loads)^2) / (sum(loads)^2 + sum(errs))
  ave[c] <- mean(loads^2)
}

rel_valid_df <- data.frame(
  Construct = constructs,
  Composite_Reliability = round(cr, 3),
  AVE = round(ave, 3)
)

ft_rel_valid <- flextable(rel_valid_df) %>%
  set_caption("Table: Composite Reliability & AVE") %>%
  autofit()

doc <- doc %>%
  body_add_par("Composite Reliability and AVE", style="heading 1") %>%
  body_add_flextable(ft_rel_valid) %>%
  body_add_par("", style="Normal")


# ===============================================================
# 6. FORNELL–LARCKER
# ===============================================================

cor_lv <- lavInspect(fit_endo_cfa, "cor.lv")[constructs, constructs]
fl_matrix <- cor_lv
diag(fl_matrix) <- sqrt(ave)

fl_df <- round(fl_matrix, 3)
fl_df <- cbind(Construct = rownames(fl_df), as.data.frame(fl_df))

ft_fl <- flextable(fl_df) %>%
  set_caption("Table: Fornell–Larcker Criterion") %>%
  autofit()

doc <- doc %>%
  body_add_par("Discriminant Validity – Fornell–Larcker", style="heading 1") %>%
  body_add_flextable(ft_fl) %>%
  body_add_par("", style="Normal")


# ===============================================================
# 7. HTMT
# ===============================================================

# (Your HTMT block stays the same)
# Assume result table is ft_htmt

doc <- doc %>%
  body_add_par("HTMT Discriminant Validity", style="heading 1") %>%
  body_add_flextable(ft_htmt) %>%
  body_add_par("", style="Normal")


# ===============================================================
# 8. STRUCTURAL PATH COEFFICIENTS
# ===============================================================

std_paths <- standardizedSolution(fit_endo) %>%
  filter(op == "~") %>%
  select(Predictor = rhs, Outcome = lhs, Beta = est.std, pvalue) %>%
  mutate(across(c(Beta, pvalue), ~round(., 3)))

ft_paths <- flextable(std_paths) %>%
  set_caption("Table: Structural Path Coefficients (β)") %>%
  autofit()

doc <- doc %>%
  body_add_par("Structural Path Coefficients", style="heading 1") %>%
  body_add_flextable(ft_paths) %>%
  body_add_par("", style="Normal")


# ===============================================================
# 9. DESCRIPTIVE STATISTICS
# ===============================================================

likert_items <- unique(loadings_df$Indicator)
likert_data <- data_used[, likert_items]

desc <- psych::describe(likert_data)[,c("mean","sd","skew","kurtosis")]
desc <- round(desc, 2)
desc <- tibble::rownames_to_column(as.data.frame(desc), "Item")

ft_desc <- flextable(desc) %>%
  set_caption("Table: Descriptive Statistics of Items") %>%
  autofit()

doc <- doc %>%
  body_add_par("Descriptive Statistics", style="heading 1") %>%
  body_add_flextable(ft_desc) %>%
  body_add_par("", style="Normal")


# ===============================================================
# 10. SAVE THE DOCUMENT
# ===============================================================

print(doc, target = "Endogenous_Model_Results.docx")


#Graphs for EV Model

# === DEFINE CONSTRUCTS FOR ENDOGENOUS MODEL ==================================
endo_constructs <- list(
  "Perceived Usefulness"      = c("PU_A", "PU_B", "PU_C", "PU_D", "PU_E"),
  "Perceived Ease of Use"     = c("PEOU_A", "PEOU_B", "PEOU_C", "PEOU_D", "PEOU_E"),
  "Willingness to Use"        = c("WTU_A", "WTU_B", "WTU_C", "WTU_D", "WTU_E"),
  "Actual Usage"              = c("AU_A",  "AU_B",  "AU_C",  "AU_D",  "AU_E")
)

# === OUTPUT DIRECTORY FOR ENDOGENOUS DESCRIPTIVES ============================
out_dir <- "C:/Users/Dominic Ocloo/Documents/Thesis_Analysis/Thesis_Analysis_Endogenous"
dir.create(out_dir, showWarnings = FALSE, recursive = TRUE)

# === LIKERT SETTINGS =========================================================
lk_levels <- c("Strongly Disagree", "Disagree", "Neutral", "Agree", "Strongly Agree")
pal <- RColorBrewer::brewer.pal(5, "RdYlGn")
names(pal) <- lk_levels

coerce_likert <- function(x) {
  if (is.factor(x)) {
    x <- forcats::fct_expand(x, lk_levels)
    x <- factor(x, levels = lk_levels, ordered = TRUE)
  } else {
    x <- suppressWarnings(as.integer(x))
    x <- factor(x, levels = 1:5, labels = lk_levels, ordered = TRUE)
  }
  levels(x) <- lk_levels
  x
}

legend_labels <- function(df) {
  df |>
    pivot_longer(everything(), values_to = "Resp") |>
    filter(!is.na(Resp)) |>
    count(Resp, name = "n") |>
    mutate(pct = round(100 * n / sum(n), 1),
           lbl = paste0(Resp, " (", pct, "%)")) |>
    {\(x) setNames(x$lbl, x$Resp)}()
}

# === PLOT SINGLE CONSTRUCT FUNCTION ==========================================
plot_construct <- function(data, vars, title_txt) {
  df <- data |> select(all_of(vars))
  df <- df[, vars, drop = FALSE]
  df <- df |> mutate(across(everything(), coerce_likert)) |> as.data.frame()
  
  lik_obj <- likert(df)
  lab_vec <- legend_labels(df)
  
  p <- plot(lik_obj, facet = FALSE,
            low.color = pal[1], high.color = pal[5]) +
    scale_fill_manual(values = pal,
                      breaks = lk_levels,
                      labels = lab_vec,
                      name = "Response") +
    labs(title = title_txt,
         y = "Proportion of Responses (%)") +
    theme_minimal(base_family = "sans") +
    theme(axis.title.x = element_blank())
  
  ggsave(file.path(out_dir,
                   paste0("Likert_", str_replace_all(title_txt, "\\s+", "_"), ".png")),
         p, width = 10, height = 6, dpi = 300)
  print(p)
}

# === SET INPUT DATA ==========================================================
input_data <- thesisdata

# === PLOT INDIVIDUAL ENDOGENOUS CONSTRUCTS ===================================
for (ct in names(endo_constructs)) {
  vars <- endo_constructs[[ct]]
  if (all(vars %in% names(input_data))) {
    plot_construct(input_data, vars, ct)
  } else {
    warning(sprintf("⚠️ Missing items in construct: %s", ct))
  }
}

# === COMBINED PLOT FOR ENDOGENOUS CONSTRUCTS =================================
combined_long <- purrr::imap_dfr(
  endo_constructs,
  function(vars, cname) {
    input_data |>
      select(all_of(vars)) |>
      mutate(across(everything(), coerce_likert)) |>
      mutate(row_id = row_number()) |>
      pivot_longer(-row_id, names_to = "Item", values_to = "Resp") |>
      mutate(Construct = cname,
             Item = factor(Item, levels = vars))
  }
) |> filter(!is.na(Resp))

summary_df <- combined_long |>
  count(Construct, Resp) |>
  group_by(Construct) |>
  mutate(Percent = 100 * n / sum(n)) |>
  ungroup()

legend_df <- summary_df |>
  group_by(Resp) |>
  summarise(avg = round(mean(Percent), 1), .groups = "drop") |>
  mutate(lbl = paste0(Resp, " (", avg, "%)"))

comb_labels <- setNames(legend_df$lbl, legend_df$Resp)
summary_df$Construct <- factor(summary_df$Construct, levels = names(endo_constructs))

p_comb <- ggplot(summary_df,
                 aes(x = Construct, y = Percent, fill = Resp)) +
  geom_bar(stat = "identity", width = .8) +
  coord_flip() +
  scale_fill_manual(values = pal,
                    breaks = lk_levels,
                    labels = comb_labels,
                    name = "Response") +
  labs(title = "Likert Distribution Across Endogenous Constructs",
       y = "Proportion of Responses (%)", x = NULL) +
  theme_minimal(base_family = "sans") +
  theme(legend.position = "right")

ggsave(file.path(out_dir, "Likert_Endogenous_Constructs_Combined.png"),
       p_comb, width = 11, height = 7, dpi = 300)
print(p_comb)

message("✅ Descriptive plots saved to: ", out_dir)


#-----------------------------------------

# -------------------------------------------------------------
# FULL CLEAN SCRIPT – CFA + SEM + ALL TABLES → WORD DOCUMENT
# -------------------------------------------------------------

library(lavaan)
library(psych)
library(flextable)
library(officer)
library(dplyr)
library(tibble)

# -------------------------------------------------------------
# 1. DEFINE MODELS
# -------------------------------------------------------------

model_external_cfa <- '
  Actual_Usage             =~ AU_A + AU_B + AU_C + AU_D + AU_E
  Availability_Infra       =~ AOI_A + AOI_B + AOI_C + AOI_D + AOI_E
  Govt_Intervention        =~ ROGI_A + ROGI_B + ROGI_C + ROGI_D + ROGI_E
  Env_Consciousness        =~ EC_A + EC_B + EC_C + EC_D + EC_E
  Perceived_Risk           =~ PR_A + PR_B + PR_C + PR_D + PR_E
  Social_Influence         =~ SCI_A + SCI_B + SCI_C + SCI_D + SCI_E
'

model_external_only <- '
  Actual_Usage             =~ AU_A + AU_B + AU_C + AU_D + AU_E
  Availability_Infra       =~ AOI_A + AOI_B + AOI_C + AOI_D + AOI_E
  Govt_Intervention        =~ ROGI_A + ROGI_B + ROGI_C + ROGI_D + ROGI_E
  Env_Consciousness        =~ EC_A + EC_B + EC_C + EC_D + EC_E
  Perceived_Risk           =~ PR_A + PR_B + PR_C + PR_D + PR_E
  Social_Influence         =~ SCI_A + SCI_B + SCI_C + SCI_D + SCI_E

  Actual_Usage ~ Availability_Infra + Govt_Intervention + Env_Consciousness +
                 Perceived_Risk + Social_Influence

'

# -------------------------------------------------------------
# 2. FIT MODELS
# -------------------------------------------------------------

fit_external_cfa <- cfa(
  model = model_external_cfa,
  data = thesisdata,
  estimator = "WLSMV",
  ordered = colnames(thesisdata)[sapply(thesisdata, is.factor)]
)

fit_external <- sem(
  model = model_external_only,
  data = thesisdata,
  estimator = "WLSMV",
  ordered = colnames(thesisdata)[sapply(thesisdata, is.factor)]
)

# -------------------------------------------------------------
# 3. CREATE WORD DOCUMENT
# -------------------------------------------------------------

doc <- read_docx()

# -------------------------------------------------------------
# SECTION A — CFA RESULTS
# -------------------------------------------------------------

# CFA Fit Indices
fit_indices_cfa <- fitMeasures(fit_external_cfa,
                               c("chisq","df","pvalue","cfi","tli","rmsea","srmr"))

fit_cfa_df <- data.frame(
  Index = c("Chi-square","DF","p-value","CFI","TLI","RMSEA","SRMR"),
  Value = round(unname(fit_indices_cfa),3)
)

ft_cfa_fit <- flextable(fit_cfa_df) %>% 
  set_caption("CFA Model Fit Indices") %>% 
  autofit()

doc <- doc %>% 
  body_add_par("CFA – Measurement Model Fit", style="heading 1") %>%
  body_add_flextable(ft_cfa_fit) %>%
  body_add_par("")

# CFA Loadings
loadings_df <- parameterEstimates(fit_external_cfa, standardized=TRUE) %>%
  filter(op=="=~") %>%
  select(Latent = lhs, Indicator = rhs, Std_Loading = std.all) %>%
  mutate(Std_Loading = round(Std_Loading,3))

ft_load <- flextable(loadings_df) %>%
  set_caption("CFA Standardized Loadings") %>% 
  autofit()

doc <- doc %>% 
  body_add_par("CFA Standardized Loadings", style="heading 1") %>%
  body_add_flextable(ft_load) %>%
  body_add_par("")

# -------------------------------------------------------------
# SECTION B — RELIABILITY (ALPHA, CR, AVE)
# -------------------------------------------------------------

latent_vars <- lavNames(fit_external_cfa, type="lv")
data_used <- lavInspect(fit_external_cfa, "data")
model_params <- parameterEstimates(fit_external_cfa) %>% filter(op=="=~")

# Safe numeric conversion
safe_num <- function(x) {
  if (is.numeric(x)) return(x)
  if (is.factor(x)) x <- as.character(x)
  as.numeric(trimws(x))
}

# Cronbach Alpha
alpha_list <- list()

for (lv in latent_vars) {
  inds <- model_params %>% filter(lhs==lv) %>% pull(rhs)
  ds_lv <- as.data.frame(data_used[,inds])
  ds_lv[] <- lapply(ds_lv, safe_num)

  ares <- psych::alpha(ds_lv)
  alpha_list[[lv]] <- data.frame(
    Construct = lv,
    Alpha = round(ares$total$raw_alpha,3)
  )
}

alpha_df <- bind_rows(alpha_list)

ft_alpha <- flextable(alpha_df) %>% 
  set_caption("Construct-Level Cronbach Alpha") %>% 
  autofit()

doc <- doc %>% 
  body_add_par("Reliability – Cronbach Alpha", style="heading 1") %>%
  body_add_flextable(ft_alpha) %>%
  body_add_par("")

# Composite Reliability & AVE
std_est <- standardizedSolution(fit_external_cfa) %>% filter(op=="=~")

constructs <- unique(std_est$lhs)
CR <- AVE <- numeric(length(constructs))
names(CR) <- names(AVE) <- constructs

for (lv in constructs) {
  load <- std_est %>% filter(lhs==lv) %>% pull(est.std)
  err <- 1 - load^2
  CR[lv] <- round((sum(load)^2)/(sum(load)^2 + sum(err)),3)
  AVE[lv] <- round(mean(load^2),3)
}

rel_df <- data.frame(
  Construct = constructs,
  Composite_Reliability = CR,
  AVE = AVE
)

ft_rel <- flextable(rel_df) %>% 
  set_caption("Composite Reliability and AVE") %>%
  autofit()

doc <- doc %>% 
  body_add_par("Composite Reliability & AVE", style="heading 1") %>%
  body_add_flextable(ft_rel) %>%
  body_add_par("")

# -------------------------------------------------------------
# SECTION C — DISCRIMINANT VALIDITY (FL, HTMT)
# -------------------------------------------------------------

# Fornell–Larcker
cor_lv <- lavInspect(fit_external_cfa, "cor.lv")[constructs,constructs]
fl_mat <- cor_lv
diag(fl_mat) <- sqrt(AVE)

fl_df <- as.data.frame(round(fl_mat,3))
fl_df$Construct <- rownames(fl_df)
fl_df <- fl_df[,c("Construct",constructs)]

ft_fl <- flextable(fl_df) %>% 
  set_caption("Fornell–Larcker Criterion") %>%
  autofit()

doc <- doc %>% 
  body_add_par("Discriminant Validity – Fornell–Larcker", style="heading 1") %>%
  body_add_flextable(ft_fl) %>%
  body_add_par("")

# HTMT
load_map <- std_est %>% select(latent = lhs, item = rhs)
items <- load_map$item

cor_items <- lavInspect(fit_external_cfa, "sampstat")$cov
cor_items <- cov2cor(cor_items)[items, items]

HTMT <- matrix(1, length(constructs), length(constructs),
               dimnames=list(constructs,constructs))

for (i in 1:length(constructs)) {
  for (j in 1:length(constructs)) {
    if (i<j) {
      it1 <- load_map %>% filter(latent==constructs[i]) %>% pull(item)
      it2 <- load_map %>% filter(latent==constructs[j]) %>% pull(item)
      HTMT[i,j] <- HTMT[j,i] <- mean(abs(cor_items[it1,it2]), na.rm=TRUE)
    }
  }
}

htmt_df <- data.frame(
  Construct = constructs,
  round(HTMT,3)
)

ft_htmt <- flextable(htmt_df) %>% 
  set_caption("HTMT Matrix") %>% 
  autofit()

doc <- doc %>% 
  body_add_par("Discriminant Validity – HTMT", style="heading 1") %>%
  body_add_flextable(ft_htmt) %>%
  body_add_par("")

# -------------------------------------------------------------
# SECTION D — ITEM-LEVEL ALPHA
# -------------------------------------------------------------

item_alpha_list <- list()

for (lv in latent_vars) {
  inds <- model_params %>% filter(lhs==lv) %>% pull(rhs)

  for (ind in inds) {
    remain <- setdiff(inds, ind)
    ds <- as.data.frame(data_used[,remain,drop=FALSE])
    ds[] <- lapply(ds, safe_num)
    ares <- psych::alpha(ds)

    item_alpha_list[[length(item_alpha_list)+1]] <-
      data.frame(Construct=lv, Item=ind,
                 Alpha_if_deleted=round(ares$total$raw_alpha,3))
  }
}

item_alpha_df <- bind_rows(item_alpha_list)

ft_item_alpha <- flextable(item_alpha_df) %>%
  set_caption("Item-Level Cronbach Alpha") %>%
  autofit()

doc <- doc %>% 
  body_add_par("Item-Level Reliability", style="heading 1") %>%
  body_add_flextable(ft_item_alpha) %>%
  body_add_par("")

# -------------------------------------------------------------
# SECTION E — SEM RESULTS
# -------------------------------------------------------------
library(lavaan)
library(dplyr)
library(flextable)
library(officer)

# --- Extract Parameter Estimates with SEs and 95% CIs ---
std_external <- param_external %>%
  filter(op == "~") %>%  # Only structural paths
  select(lhs, rhs, std.all, se, pvalue, ci.lower, ci.upper) %>%
  rename(
    Outcome = lhs,
    Predictor = rhs,
    `Standardized Estimate (β)` = std.all,
    `Standard Error (SE)` = se,
    `p-value` = pvalue,
    `95% CI Lower` = ci.lower,
    `95% CI Upper` = ci.upper
  ) %>%
  mutate(across(`Standardized Estimate (β)`:`95% CI Upper`, ~round(., 3)))

# --- Format Flextable for Word ---
ft_external <- flextable(std_external) %>%
  set_caption("Table: Structural Path Coefficients – External Predictors of Actual Usage (with SEs and 95% CIs)") %>%
  autofit() %>%
  fontsize(size = 12, part = "all") %>%
  font(fontname = "Times New Roman", part = "all") %>%
  flextable::align(align = "center", part = "all") %>%
  border_remove() %>%
  border_outer(border = fp_border(color = "black")) %>%
  border_inner_h(border = fp_border(color = "black")) %>%
  border_inner_v(border = fp_border(color = "black"))

# --- Add heading and table to the Word document ---
doc <- body_add_par(doc, "External Factors Model Results", style = "heading 1")
doc <- body_add_flextable(doc, ft_external)
doc <- body_add_par(doc, "", style = "Normal")  # Line space


# -----------------------------------------------------------------
# --- Descriptive Statistics – Likert‑Scale Items (external model) 
#------------------------------------------------------------------

library(dplyr)
library(psych)
library(flextable)
library(lavaan)
library(tibble)

# 1. Pull the raw data actually analysed (use thesisdata directly)
likert_items <- parameterEstimates(fit_external) %>%          # or fit_external_clean
  filter(op == "=~") %>%                                      # measurement relations
  pull(rhs) %>% 
  unique()

likert_data <- thesisdata[ , likert_items, drop = FALSE]       # original columns

# 2. Robust numeric conversion --------------------------------------------------
verbal_levels  <- c("Strongly Disagree", "Disagree", "Neutral",
                    "Agree", "Strongly Agree")                 # adjust if needed
numeric_levels <- as.character(1:5)

likert_num <- data.frame(
  lapply(names(likert_data), function(var) {
    x <- likert_data[[var]]
    
    # already numeric
    if (is.numeric(x)) return(x)
    
    # factor with numeric levels ("1"…"5")
    if (is.factor(x) && all(levels(x) %in% numeric_levels)) {
      return(as.numeric(as.character(x)))
    }
    
    # text / factor with verbal labels
    x_chr <- trimws(as.character(x))
    if (all(x_chr %in% verbal_levels | is.na(x_chr))) {
      return(match(x_chr, verbal_levels))                      # 1…5
    }
    
    # unknown coding → NA + message
    message("Item '", var, "' has unexpected labels; recoded to NA.")
    return(rep(NA_real_, length(x)))
  }),
  check.names = FALSE
)
names(likert_num) <- likert_items     # restore column names

# 3. Drop items with all NA or no variance --------------------------------------
keep <- sapply(names(likert_num), function(var) {
  v <- likert_num[[var]]
  good <- length(unique(na.omit(v))) > 1
  if (!good) message("Dropped constant/empty item: ", var)
  good
})
likert_num <- likert_num[ , keep, drop = FALSE]

if (ncol(likert_num) == 0) stop("No valid numeric Likert items found.")

# 4. Compute descriptive statistics --------------------------------------------
desc_stats <- psych::describe(likert_num)[ , c("n", "mean", "sd", "skew", "kurtosis")]
desc_stats <- round(desc_stats, 2) %>%
  as.data.frame() %>%
  rownames_to_column("Item")

# 5. Create flextable -----------------------------------------------------------
ft_desc <- flextable(desc_stats) %>%
  set_caption("Table X: Descriptive Statistics of Likert‑Scale Items") %>%
  autofit() %>%
  fontsize(size = 12, part = "all") %>%
  font(fontname = "Times New Roman", part = "all") %>%
  flextable::align(align = "center", part = "all") %>%
  border_remove() %>%
  border_outer(border = fp_border(color = "black")) %>%
  border_inner_h(border = fp_border(color = "black")) %>%
  border_inner_v(border = fp_border(color = "black"))

# 6. Append the table to the Word document --------------------------------------
doc <- body_add_par(doc,
                    "Descriptive Statistics for Likert‑Scale Items (External Model)",
                    style = "heading 1")
doc <- body_add_flextable(doc, ft_desc)
doc <- body_add_par(doc, "", style = "Normal")

# -------------------------------------------------------------
# SECTION E — SEM RESULTS (External Predictors of Actual Usage)
# -------------------------------------------------------------

library(lavaan)
library(dplyr)
library(flextable)
library(officer)

# --- Extract Parameter Estimates with Standard Errors and 95% CIs ---
param_external <- parameterEstimates(fit_external, standardized = TRUE, ci = TRUE)

std_external <- param_external %>%
  filter(op == "~") %>%  # Only structural paths
  select(lhs, rhs, std.all, se, pvalue, ci.lower, ci.upper) %>%
  rename(
    Outcome = lhs,
    Predictor = rhs,
    `Standardized Estimate (β)` = std.all,
    `Standard Error (SE)` = se,
    `p-value` = pvalue,
    `95% CI Lower` = ci.lower,
    `95% CI Upper` = ci.upper
  ) %>%
  mutate(across(`Standardized Estimate (β)`:`95% CI Upper`, ~round(., 3)))  # Round numeric columns

# --- Format Flextable ---
ft_external <- flextable(std_external) %>%
  set_caption("Table: Structural Path Coefficients – External Predictors of Actual Usage (with SEs and 95% CIs)") %>%
  autofit() %>%
  fontsize(size = 12, part = "all") %>%
  font(fontname = "Times New Roman", part = "all") %>%
  flextable::align(align = "center", part = "all") %>%
  border_remove() %>%
  border_outer(border = fp_border(color = "black")) %>%
  border_inner_h(border = fp_border(color = "black")) %>%
  border_inner_v(border = fp_border(color = "black"))

# --- Add heading and table to the Word document ---
doc <- doc %>%
  body_add_par("External Factors Model Results", style = "heading 1") %>%
  body_add_flextable(ft_external) %>%
  body_add_par("", style = "Normal")  # Line space


# -------------------------------------------------------------
# SAVE WORD DOCUMENT
# -------------------------------------------------------------

print(doc, target="External_Model_Full_Report.docx")


# Load the required libraries
library(officer)
library(flextable)

# Step 1: Extract R-squared values
r_squared_values <- inspect(fit_external, "r2")

# Step 2: Convert to a data frame
r2_df <- data.frame(
  Variable = names(r_squared_values),
  R_Squared = round(as.numeric(r_squared_values), 3)
)

# Step 3: Create and export Word document
doc <- read_docx() %>%
  body_add_par("Table: R-squared Values from SEM Model (fit_external)", style = "heading 1") %>%
  body_add_flextable(flextable(r2_df))

# Save the Word document
print(doc, target = "EVR_squared_fit_external.docx")


#Visualize Modle with Loadings

# Generate semPlotModel object from fitted model
plot_obj <- semPlotModel(fit_external)

# Extract all node names
node_names <- plot_obj@Vars$name
# Extract whether each node is manifest (TRUE) or latent (FALSE)
node_is_manifest <- plot_obj@Vars$manifest

# Get latent variable names
latent_vars <- node_names[!node_is_manifest]

# Function to convert variable name to initials
get_initials <- function(name) {
  # Split by underscores or spaces
  parts <- unlist(strsplit(name, split = "_|\\s"))
  # Take first letter of each part and combine
  paste(toupper(substr(parts, 1, 1)), collapse = "")
}

# Create a named vector mapping latent variable names to their initials
label_map <- setNames(sapply(latent_vars, get_initials), latent_vars)

# Now create labels for all nodes
custom_labels <- sapply(node_names, function(x) {
  if (!node_is_manifest[node_names == x] && x %in% names(label_map)) {
    label_map[[x]]   # Use initials for latent variables
  } else {
    x               # Keep original for manifest variables (items)
  }
})

# Assign colors as before
node_colors <- rep(NA, length(node_names))
names(node_colors) <- node_names

for (i in seq_along(node_names)) {
  if (!node_is_manifest[i]) {
    if (node_names[i] == "Actual_Usage") {
      node_colors[i] <- "darkgreen"
    } else {
      node_colors[i] <- "orange"
    }
  } else {
    node_colors[i] <- "darkviolet"
  }
}

# Plot the model
semPaths(
  object = plot_obj,
  whatLabels = "std",
  style = "lisrel",
  layout = "tree",
  nodeLabels = custom_labels,
  edge.label.cex = 0.9,
  sizeLat = 7,
  sizeMan = 5,
  color = node_colors,
  mar = c(6, 6, 6, 6),
  intercepts = FALSE,
  residuals = FALSE,
  optimizeLatRes = TRUE,
  edge.color = "black"
)


#Visualize Model Without Item loadings

# Generate semPlotModel object from fitted model
plot_obj <- semPlotModel(fit_external)

# Extract all node names and manifest status
node_names <- plot_obj@Vars$name
node_is_manifest <- plot_obj@Vars$manifest

# Function to get initials from latent variable name
get_initials <- function(name) {
  parts <- unlist(strsplit(name, split = "_|\\s"))
  paste(toupper(substr(parts, 1, 1)), collapse = "")
}

# Get latent variable names
latent_vars <- node_names[!node_is_manifest]

# Create label map dynamically: latent variable name -> initials
label_map <- setNames(sapply(latent_vars, get_initials), latent_vars)

# Create custom labels: initials for latent variables, empty string for manifest
custom_labels <- sapply(node_names, function(x) {
  if (!node_is_manifest[node_names == x] && x %in% names(label_map)) {
    label_map[[x]]
  } else {
    ""  # Hide manifest variable labels
  }
})

# Set node colors: latent variables orange or darkgreen, manifest variables transparent
node_colors <- ifelse(
  node_is_manifest, 
  "transparent",             # hide manifest variable nodes
  ifelse(node_names == "Actual_Usage", "darkgreen", "orange")
)

# Plot with manifest variables hidden (sizeMan = 0)
semPaths(
  object = plot_obj,
  whatLabels = "std",
  style = "lisrel",
  layout = "tree",
  nodeLabels = custom_labels,
  edge.label.cex = 0.9,
  sizeLat = 7,
  sizeMan = 0,               # hide manifest nodes
  color = node_colors,
  mar = c(6, 6, 6, 6),
  intercepts = FALSE,
  residuals = FALSE,
  optimizeLatRes = TRUE,
  edge.color = "black",
  curvePivot = 0.5           # Increase the distance/offset of curved edges (default is 0-3)
)

#--------------------------------------------


#Visualization of Likert Plot

# === DEFINE CONSTRUCTS ========================================================
constructs <- list(
  "Actual Usage"             = c("AU_A", "AU_B", "AU_C", "AU_D", "AU_E"),
  "Availability of Infrastructure" = c("AOI_A", "AOI_B", "AOI_C", "AOI_D", "AOI_E"),
  "Government Intervention"  = c("ROGI_A", "ROGI_B", "ROGI_C", "ROGI_D", "ROGI_E"),
  "Environmental Consciousness" = c("EC_A", "EC_B", "EC_C", "EC_D", "EC_E"),
  "Perceived Risk"           = c("PR_A", "PR_B", "PR_C", "PR_D", "PR_E"),
  "Socio-Cultural Influence"         = c("SCI_A", "SCI_B", "SCI_C", "SCI_D", "SCI_E")
)

# === OUTPUT DIRECTORY =========================================================
out_dir <- "C:/Users/Dominic Ocloo/Documents/Thesis_Analysis/Thesis_Analysis_External"
dir.create(out_dir, showWarnings = FALSE, recursive = TRUE)

# === LIKERT LEVELS AND PALETTE ================================================
lk_levels <- c("Strongly Disagree", "Disagree", "Neutral", "Agree", "Strongly Agree")
pal <- RColorBrewer::brewer.pal(5, "RdYlGn")
names(pal) <- lk_levels

# === LIKERT COERCION FUNCTION =================================================
coerce_likert <- function(x) {
  if (is.factor(x)) {
    x <- forcats::fct_expand(x, lk_levels)
    x <- factor(x, levels = lk_levels, ordered = TRUE)
  } else {
    x <- suppressWarnings(as.integer(x))
    x <- factor(x, levels = 1:5, labels = lk_levels, ordered = TRUE)
  }
  levels(x) <- lk_levels
  x
}

# === LEGEND LABEL GENERATOR ===================================================
legend_labels <- function(df) {
  df |>
    pivot_longer(everything(), values_to = "Resp") |>
    filter(!is.na(Resp)) |>
    count(Resp, name = "n") |>
    mutate(pct = round(100 * n / sum(n), 1),
           lbl = paste0(Resp, " (", pct, "%)")) |>
    {\(x) setNames(x$lbl, x$Resp)}()
}

# === PLOT FUNCTION FOR SINGLE CONSTRUCT =======================================
plot_construct <- function(data, vars, title_txt) {
  df <- data |>
    select(all_of(vars)) |>
    mutate(across(everything(), coerce_likert)) |>
    as.data.frame()
  
  lik_obj <- likert(df)
  lab_vec <- legend_labels(df)
  
  p <- plot(lik_obj, facet = FALSE,
            low.color = pal[1], high.color = pal[5]) +
    scale_fill_manual(values = pal,
                      breaks = lk_levels,
                      labels = lab_vec,
                      name = "Response") +
    labs(title = title_txt,
         y = "Proportion of Responses (%)") +
    theme_minimal(base_family = "sans") +
    theme(axis.title.x = element_blank())
  
  ggsave(file.path(out_dir,
                   paste0("Likert_", str_replace_all(title_txt, "\\s+", "_"), ".png")),
         p, width = 10, height = 6, dpi = 300)
  print(p)
}

# === SET DATA =================================================================
input_data <- thesisdata

# === PLOT INDIVIDUAL CONSTRUCTS ==============================================
for (ct in names(constructs)) {
  vars <- constructs[[ct]]
  if (all(vars %in% names(input_data))) {
    plot_construct(input_data, vars, ct)
  } else {
    warning(sprintf("Missing items in construct: %s", ct))
  }
}

# === COMBINED PLOT FOR ALL CONSTRUCTS ========================================
combined_long <- purrr::imap_dfr(
  constructs,
  \(vars, cname) {
    input_data |>
      select(all_of(vars)) |>
      mutate(across(everything(), coerce_likert)) |>
      mutate(row_id = row_number()) |>
      pivot_longer(-row_id, names_to = "Item", values_to = "Resp") |>
      mutate(Construct = cname)
  }
) |>
  filter(!is.na(Resp))

# === SUMMARY FOR COMBINED =====================================================
summary_df <- combined_long |>
  count(Construct, Resp) |>
  group_by(Construct) |>
  mutate(Percent = 100 * n / sum(n)) |>
  ungroup()

legend_df <- summary_df |>
  group_by(Resp) |>
  summarise(avg = round(mean(Percent), 1), .groups = "drop") |>
  mutate(lbl = paste0(Resp, " (", avg, "%)"))

comb_labels <- setNames(legend_df$lbl, legend_df$Resp)
summary_df$Construct <- factor(summary_df$Construct, levels = names(constructs))

# === FINAL COMBINED PLOT ======================================================
p_comb <- ggplot(summary_df,
                 aes(x = Construct, y = Percent, fill = Resp)) +
  geom_bar(stat = "identity", width = .8) +
  coord_flip() +
  scale_fill_manual(values = pal,
                    breaks = lk_levels,
                    labels = comb_labels,
                    name = "Response") +
  labs(title = "Likert Distribution Across External Constructs",
       y = "Proportion of Responses (%)", x = NULL) +
  theme_minimal(base_family = "sans") +
  theme(legend.position = "right")

ggsave(file.path(out_dir, "Likert_External_Constructs_Combined.png"),
       p_comb, width = 11, height = 7, dpi = 300)
print(p_comb)

#Check Folder holding Analysis

