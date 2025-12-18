# Convert "Location within Accra Metropolis" to a factor
thesisdata$`location_within_accra_metropolis` <- factor(
  thesisdata$`location_within_accra_metropolis`, 
  levels = c("Ablekuma South Sub-Metropolitan", 
             "Okaikwei South Sub-Metropolitan", 
             "Ashiedu Keteke Sub-Metropolitan")
)

# Convert "Sex" to a factor
thesisdata$sex <- factor(thesisdata$sex, levels = c("Male", "Female"))

#Convert "Marital status" to a factor
thesisdata$`marital_status` <- factor(thesisdata$`marital_status`,
                                     levels = c("Single",
                                                "Married",
                                                "Divorced",
                                                "Widowed")
)
# Load the dplyr package
library(dplyr)

# Convert monthly income to numeric midpoint
# Clean and reassign midpoint to income
thesisdata$monthly_income <- case_when(
  thesisdata$monthly_income == "Below GH₵1,500" ~ 1000,
  thesisdata$monthly_income == "Between GH₵1,500 and GH₵5,000" ~ 3250,
  thesisdata$monthly_income == "Between GH₵5,000 and GH₵10,000" ~ 7500,
  thesisdata$monthly_income == "Above GH₵10,000" ~ 12500,
  TRUE ~ NA_real_
)

# View updated column
head(thesisdata$monthly_income)


# Convert "Level of education completed" to an ordered factor
thesisdata$`level_of_education_completed` <- factor(
  thesisdata$`level_of_education_completed`, 
  levels = c("Basic Education", 
             "Pre-tertiary Education", 
             "Tertiary Education"),
  ordered = TRUE
)

# Convert "Type of vehicle currently own or driven" to a factor
thesisdata$`type_of_vehicle_currently_own_or_driven` <- factor(
  thesisdata$`type_of_vehicle_currently_own_or_driven`, 
  levels = c("Private vehicle", "Commercial vehicle")
)

# Convert "Do you own the vehicle you drive currently?" to a factor
thesisdata$`do_you_own_the_vehicle_you_drive_currently` <- factor(
  thesisdata$`do_you_own_the_vehicle_you_drive_currently`, 
  levels = c("Yes", "No", "Partly owned")
)

# Convert "Do you support the conversion or Retrofitting of conventional vehicles into Electric Vehicles?" to a factor
thesisdata$`do_you_support_the_conversion_or_retrofitting_of_conventional_vehicles_into_electric_vehicles` <- factor(
  thesisdata$`do_you_support_the_conversion_or_retrofitting_of_conventional_vehicles_into_electric_vehicles`, 
  levels = c("Yes", "No", "Indifferent")
)

# Convert the column to a factor
thesisdata$`do_you_think_government_is_doing_enough_to_support_electric_vehicle_adoption` <- factor(
  thesisdata$`do_you_think_government_is_doing_enough_to_support_electric_vehicle_adoption`, 
  levels = c("Yes", "No", "Indifferent")
)

# Load necessary library
library(dplyr)

# Convert "Type of driver" to a factor
thesisdata$`type_of_driver` <- factor(
  thesisdata$`type_of_driver`, 
  levels = c("Internal Combustion Engine Vehicle driver (Patrol, diesel etc. engine vehicle)", 
             "Electric Vehicle Driver")
)

thesisdata$type_of_driver_clean <- case_when(
  grepl("Internal Combustion Engine Vehicle driver", thesisdata$type_of_driver, ignore.case = TRUE) ~ "ICE",
  grepl("Electric Vehicle Driver", thesisdata$type_of_driver, ignore.case = TRUE) ~ "EV",
  TRUE ~ NA_character_
)

ice_data <- filter(thesisdata, type_of_driver_clean == "ICE")
ev_data <- filter(thesisdata, type_of_driver_clean == "EV")


# Define Sociodemographic Questions (applies to all respondents)
sociodemographic_questions <- c(
  "location_within_accra_metropolis", "sex", "marital_status", "age", "occupation", 
  "years_of_driving_experience", "monthly_income", "level_of_education_completed", 
  "type_of_vehicle_currently_own_or_driven", "do_you_own_the_vehicle_you_drive_currently", 
  "how_many_kilometers_do_you_cover_daily", 
  "do_you_support_the_conversion_or_retrofitting_of_conventional_vehicles_into_electric_vehicles", 
  "do_you_think_government_is_doing_enough_to_support_electric_vehicle_adoption"
)

# Define skip logic using metadata
attributes(thesisdata)$skip_logic <- list(
  
  # Internal Combustion Engine (ICE) Driver Questions
  ICE_Driver = c(
    "have_you_heard_of_cars_that_run_solely_on_battery_power_without_a_fuel_engine",
    "are_you_likely_to_buy_another_car_in_the_next_20_years_or_sooner",
    "do_you_have_plans_of_buying_an_electric_vehicles_in_the_next_20_years_or_sooner",
    "PU_1", "PU_2", "PU_3", "PU_4", "PU_5",
    "PEOU_1", "PEOU_2", "PEOU_3", "PEOU_4", "PEOU_5",
    "WTU_1", "WTU_2", "WTU_3", "WTU_4", "WTU_5",
    "AOI_1", "AOI_2", "AOI_3", "AOI_4", "AOI_5",
    "ROGI_1", "ROGI_2", "ROGI_3", "ROGI_4", "ROGI_5",
    "EC_1", "EC_2", "EC_3", "EC_4", "EC_5",
    "PR_1", "PR_2", "PR_3", "PR_4", "PR_5",
    "SCI_1", "SCI_2", "SCI_3", "SCI_4", "SCI_5",
    "Comments"
  ),
  
  # Electric Vehicle (EV) Driver Questions
  EV_Driver = c(
    "what_type_of_ev_do_you_drive", 
    "how_long_have_you_owned_or_used_your_ev",
    "PU_A", "PU_B", "PU_C", "PU_D", "PU_E",
    "PEOU_A", "PEOU_B", "PEOU_C", "PEOU_D", "PEOU_E",
    "WTU_A", "WTU_B", "WTU_C", "WTU_D", "WTU_E",
    "AU_A", "AU_B", "AU_C", "AU_D", "AU_E",
    "AOI_A", "AOI_B", "AOI_C", "AOI_D", "AOI_E",
    "ROGI_A", "ROGI_B", "ROGI_C", "ROGI_D", "ROGI_E",
    "EC_A", "EC_B", "EC_C", "EC_D", "EC_E",
    "PR_A", "PR_B", "PR_C", "PR_D", "PR_E",
    "SCI_A", "SCI_B", "SCI_C", "SCI_D", "SCI_E",
    "Comments"
  )
)

# Function to filter dataset based on driver type (while keeping sociodemographic questions)
filter_questions <- function(driver_type) {
  # Print the driver type to verify correctness
  print(paste("Filtering for driver type:", driver_type))
  
  # Ensure driver_type is correctly assigned
  if (driver_type == "Internal Combustion Engine Vehicle driver (Patrol, diesel etc. engine vehicle)") {
    relevant_cols <- attributes(thesisdata)$skip_logic$ICE_Driver
  } else if (driver_type == "Electric Vehicle Driver") {
    relevant_cols <- attributes(thesisdata)$skip_logic$EV_Driver
  } else {
    stop("Invalid driver type. Choose 'Internal Combustion Engine Vehicle driver' or 'Electric Vehicle Driver'.")
  }
  
  # Combine relevant columns with sociodemographic questions
  selected_cols <- unique(c("Type of driver", sociodemographic_questions, relevant_cols))
  
  # Keep only columns that exist in thesisdata
  existing_cols <- selected_cols[selected_cols %in% colnames(thesisdata)]
  
  # Print selected columns for debugging
  print("Selected Columns:")
  print(existing_cols)
  
  # Subset dataset for the selected driver type
  return(thesisdata %>% select(all_of(existing_cols)))
}

# Example Usage:
# Get dataset for Internal Combustion Engine Vehicle Drivers
ice_data <- filter_questions("Internal Combustion Engine Vehicle driver (Patrol, diesel etc. engine vehicle)")

# Get dataset for Electric Vehicle Drivers
ev_data <- filter_questions("Electric Vehicle Driver")

# Verify structure
str(ice_data)
str(ev_data)

# Convert the column to a factor with specified categories
thesisdata$`have_you_heard_of_cars_that_run_solely_on_battery_power_without_a_fuel_engine` <- factor(
  thesisdata$`have_you_heard_of_cars_that_run_solely_on_battery_power_without_a_fuel_engine`,
  levels = c("Yes", "No", "Indifferent")
)

# Verify the changes
str(thesisdata$`have_you_heard_of_cars_that_run_solely_on_battery_power_without_a_fuel_engine`)

# Convert the column to a factor with specified categories
thesisdata$`are_you_likely_to_buy_another_car_in_the_next_20_years_or_sooner` <- factor(
  thesisdata$`are_you_likely_to_buy_another_car_in_the_next_20_years_or_sooner`,
  levels = c("Yes", "No", "Indifferent")
)

# Verify the changes
str(thesisdata$`are_you_likely_to_buy_another_car_in_the_next_20_years_or_sooner`)

# Convert the column Do you have plans of buying an Electric Vehicles in the next 20 years or sooner? to a factor with specified categories
thesisdata$`do_you_have_plans_of_buying_an_electric_vehicles_in_the_next_20_years_or_sooner` <- factor(
  thesisdata$`do_you_have_plans_of_buying_an_electric_vehicles_in_the_next_20_years_or_sooner`,
  levels = c("Yes", "No", "Indifferent")
)

# Verify the changes
str(thesisdata$`do_you_have_plans_of_buying_an_electric_vehicles_in_the_next_20_years_or_sooner`)

# Convert the column to a factor with specified categories
thesisdata$`what_type_of_ev_do_you_drive` <- factor(
  thesisdata$`what_type_of_ev_do_you_drive`,
  levels = c("Fully/Battery Electric Vehicle", 
             "Plug-in Hybrid Electric Vehicle", 
             "Hybrid-Electric Vehicle")
)

# Verify the changes
str(thesisdata$`what_type_of_ev_do_you_drive`)

# Load necessary libraries
library(dplyr)
library(stringr)

# Define the Likert scale mapping (Text to Numeric)
likert_mapping <- c("Strongly Disagree" = 1, 
                    "Disagree" = 2, 
                    "Neutral" = 3, 
                    "Agree" = 4, 
                    "Strongly Agree" = 5)

# List of Likert scale columns (Group 1)
likert_columns_1 <- c("PU_1", "PU_2", "PU_3", "PU_4", "PU_5", 
                      "PEOU_1", "PEOU_2", "PEOU_3", "PEOU_4", "PEOU_5",
                      "WTU_1", "WTU_2", "WTU_3", "WTU_4", "WTU_5",
                      "AOI_1", "AOI_2", "AOI_3", "AOI_4", "AOI_5",
                      "ROGI_1", "ROGI_2", "ROGI_3", "ROGI_4", "ROGI_5",
                      "EC_1", "EC_2", "EC_3", "EC_4", "EC_5",
                      "PR_1", "PR_2", "PR_3", "PR_4", "PR_5",
                      "SCI_1", "SCI_2", "SCI_3", "SCI_4", "SCI_5")

# List of Likert scale columns (Group 2)
likert_columns_2 <- c("PU_A", "PU_B", "PU_C", "PU_D", "PU_E", 
                      "PEOU_A", "PEOU_B", "PEOU_C", "PEOU_D", "PEOU_E",
                      "WTU_A", "WTU_B", "WTU_C", "WTU_D", "WTU_E",
                      "AU_A", "AU_B", "AU_C", "AU_D", "AU_E",
                      "AOI_A", "AOI_B", "AOI_C", "AOI_D", "AOI_E",
                      "ROGI_A", "ROGI_B", "ROGI_C", "ROGI_D", "ROGI_E",
                      "EC_A", "EC_B", "EC_C", "EC_D", "EC_E",
                      "PR_A", "PR_B", "PR_C", "PR_D", "PR_E",
                      "SCI_A", "SCI_B", "SCI_C", "SCI_D", "SCI_E")

# Function to convert Likert responses to numeric (Text -> Numbers)
convert_to_numeric <- function(x) {
  # Remove extra spaces and convert the Likert scale responses to numbers
  as.numeric(recode(x, !!!likert_mapping))  # Recode using the likert_mapping
}

# Apply conversion to both sets of Likert scale columns (Convert to Numeric)
thesisdata[likert_columns_1] <- lapply(thesisdata[likert_columns_1], convert_to_numeric)
thesisdata[likert_columns_2] <- lapply(thesisdata[likert_columns_2], convert_to_numeric)

# Function to convert numeric values to factors (Numbers -> Factors)
convert_to_factor <- function(x) {
  factor(x, levels = c(1, 2, 3, 4, 5), 
         labels = c("Strongly Disagree", "Disagree", "Neutral", "Agree", "Strongly Agree"))
}

# Apply conversion to both sets of Likert scale columns (Convert to Factors)
thesisdata[likert_columns_1] <- lapply(thesisdata[likert_columns_1], convert_to_factor)
thesisdata[likert_columns_2] <- lapply(thesisdata[likert_columns_2], convert_to_factor)

# Verify by checking one column
table(thesisdata$PU_1, useNA = "ifany")  # Replace PU_1 with another column if needed
table(thesisdata$PU_A, useNA = "ifany")


write.csv(thesisdata, "thesisdata.csv", row.names = FALSE)

