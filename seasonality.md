---
title: "Seasonality"
order: 3
weight: 3
format:
  html:
    toc: true
    toc-depth: 4
---

## Overview

Malaria transmission follows clear seasonal trends driven by environmental factors. This analysis focuses on rainfall-based seasonality to determine the most effective periods for malaria interventions. Using Sierra Leone as a case study, the **WHO-recommended threshold** was applied to classify areas suitable for seasonal malaria control strategies. The results provide evidence to guide the planning and timing of interventions such as Seasonal Malaria Chemoprevention (SMC), helping health programs optimize resources and maximize impact.

## Understanding Seasonal Transmission Patterns

Malaria transmission varies throughout the year, driven by seasonal changes in rainfall, temperature, and humidity. These environmental factors directly influence mosquito breeding cycles and parasite development rates. Recognizing these temporal patterns is essential for optimizing malaria control strategies.

## The Importance of Seasonality

In malaria-endemic regions, cases often cluster during specific months when conditions favor transmission. By identifying these high-risk periods, public health programs can strategically time interventions to maximize their effectiveness and efficiently allocate limited resources.

## Defining Seasonality

**For this analysis, seasonality refers to transmission patterns where a significant portion of annual malaria burden occurs within a concentrated timeframe.**

Following **WHO recommendations**, a region exhibits seasonal transmission when:

- Any **four consecutive months** account for **60% or more** of the total annual malaria burden or rainfall
- This pattern remains consistent across multiple years

Areas meeting these criteria are considered suitable for targeted seasonal interventions, such as Seasonal Malaria Chemoprevention (SMC).

## Analytical Approach

This analysis utilizes **CHIRPS satellite-derived rainfall data** as a proxy for malaria transmission intensity. Sierra Leone serves as the case study example.

**Data Source**: Rainfall datasets can be obtained following the procedures outlined in the [CHIRPS data extraction guide](https://ahadi-analytics.github.io/snt-code-library/english/library/data/climate/extract_raster_climate.html).

## Profiling for SMC

![Profiling seasonality for SMC targeting](https://github.com/mohamedsillahkanu/snt-package/raw/7aca50c7688bbc679b33f5543150078c978c7edf/seasonality_eligibility_v2.png)

::: {.callout-note title="Objectives" appearance="simple"}
- Determine which districts demonstrate seasonal transmission patterns suitable for SMC implementation  
- Map the timing of rainfall and malaria transmission peaks to optimize intervention scheduling for maximum impact
:::

## Step-by-Step

### Step 1: Import packages and data

```{R}
#| message: false
#| echo: true
#| eval: true
#| code-fold: false
#| code-summary: "Show the code"
if (!requireNamespace("pacman", quietly = TRUE)) install.packages("pacman")

pacman::p_load(
  readxl, dplyr, openxlsx, lubridate, ggplot2, readr, stringr, 
  here, tidyr, gridExtra, knitr, writexl, sf
)

# File path
file_path <- here::here("english/data_r/modeled", "chirps_data_2015_2023_lastest.xls")

# Load your data
data <- readxl::read_excel(file_path)
```


**To adapt the code:**

- **Line 9**: Update the file path to correctly reference the current dataset


### Step 2: Configure analysis parameters

```{R}
#| message: false
#| echo: true
#| eval: true
#| code-fold: false
#| code-summary: "Show the code"
# Column names required in the dataset
year_column <- "Year"
month_column <- "Month"
value_column <- "mean_rain"
admin_columns <- c("FIRST_DNAM", "FIRST_CHIE")

# Analysis parameters
analysis_start_year <- 2015
analysis_start_month <- 1
seasonality_threshold <- 60

# Output file names
#detailed_output <- "detailed_seasonality_results_09_22_2025.xlsx"
#yearly_output <- "yearly_analysis_summary.xlsx"
#location_output <- "location_seasonality_summary.xlsx"
```

**To adapt the code:**

- **Lines 2-5**: Update column names to match the dataset

- **Lines 8-10**: Adjust analysis parameters such as start year, start month, and threshold based on the analysis context

- **Lines 13-15**: Modify output file names as needed for saving the results



### Step 3: Prepare data for analysis

```{R}
#| message: false
#| echo: true
#| eval: true
#| code-fold: false
#| code-summary: "Show the code"
# Check if required columns exist
required_cols <- c(year_column, month_column, value_column, admin_columns)
missing_cols <- required_cols[!required_cols %in% colnames(data)]
if (length(missing_cols) > 0) {
  stop(paste("Missing columns:", paste(missing_cols, collapse = ", ")))
}

# Filter data and ensure month is numeric
filtered_data <- data |>
  dplyr::filter(
    !is.na(!!sym(year_column)) & 
    !is.na(!!sym(month_column)) & 
    !!sym(year_column) >= analysis_start_year
  ) |>
  dplyr::mutate(Month = as.numeric(!!sym(month_column)))

# Create combined grouping variable from multiple admin unit columns
if (length(admin_columns) == 1) {
  filtered_data$admin_group <- filtered_data[[admin_columns[1]]]
} else {
  filtered_data <- filtered_data |>
    dplyr::mutate(
      admin_group = paste(!!!syms(admin_columns), sep = " | ")
    )
}

# Calculate data span and validate minimum requirements
available_years <- sort(unique(filtered_data[[year_column]]))
data_span_years <- length(available_years)

if (data_span_years < 6) {
  stop(paste(
    "Insufficient data: Analysis requires at least 6 years of data, but only", 
    data_span_years, "years found."
  ))
}

num_complete_years <- data_span_years - 1
total_num_blocks <- num_complete_years * 12
```

**To adapt the code:**

- Do not change anything in the code 


### Step 4: Generate rolling time blocks

```{R}
#| message: false
#| echo: true
#| eval: true
#| code-fold: false
#| code-summary: "Show the code"
# Create month names for labels
month_names <- c("Jan", "Feb", "Mar", "Apr", "May", "Jun",
                 "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

# Generate all time blocks
blocks <- data.frame()
current_year <- analysis_start_year
current_month <- analysis_start_month

for (i in 1:total_num_blocks) {
  # 4-month period
  start_4m_year <- current_year
  start_4m_month <- current_month
  end_4m_year <- current_year
  end_4m_month <- current_month + 3
  
  if (end_4m_month > 12) {
    end_4m_year <- end_4m_year + 1
    end_4m_month <- end_4m_month - 12
  }
  
  # 12-month period  
  start_12m_year <- current_year
  start_12m_month <- current_month
  end_12m_year <- current_year
  end_12m_month <- current_month + 11
  
  if (end_12m_month > 12) {
    end_12m_year <- end_12m_year + 1
    end_12m_month <- end_12m_month - 12
  }
  
  # Create date range label
  date_range <- paste0(
    month_names[start_4m_month], " ", start_4m_year, "-",
    month_names[end_4m_month], " ", end_4m_year
  )
  
  # Store block information
  blocks <- rbind(blocks, data.frame(
    block_number = i,
    start_4m_year = start_4m_year,
    start_4m_month = start_4m_month,
    end_4m_year = end_4m_year,
    end_4m_month = end_4m_month,
    start_12m_year = start_12m_year,
    start_12m_month = start_12m_month,
    end_12m_year = end_12m_year,
    end_12m_month = end_12m_month,
    date_range = date_range
  ))
  
  # Move to next month
  current_month <- current_month + 1
  if (current_month > 12) {
    current_month <- 1
    current_year <- current_year + 1
  }
}

```

**To adapt the code:**

- Do not change anything in the code 

### Step 5: Calculate seasonality for each block

```{R}
#| message: false
#| echo: true
#| eval: false
#| code-fold: false
#| code-summary: "Show the code"
# Initialize results dataframe
detailed_results <- data.frame()

# Get unique administrative groups
admin_groups <- unique(filtered_data$admin_group)

# Loop through each administrative unit
for (admin_unit in admin_groups) {
  # Filter data for this administrative unit
  unit_data <- filtered_data |>
    dplyr::filter(admin_group == admin_unit)
  
  # Loop through each time block
  for (i in 1:nrow(blocks)) {
    block <- blocks[i, ]
    
    # Create year-month values for comparison
    unit_data_ym <- unit_data[[year_column]] * 12 + unit_data$Month
    
    # Calculate 4-month window
    start_4m_ym <- block$start_4m_year * 12 + block$start_4m_month
    end_4m_ym <- block$end_4m_year * 12 + block$end_4m_month
    
    data_4m <- unit_data |>
      dplyr::filter(
        unit_data_ym >= start_4m_ym & unit_data_ym <= end_4m_ym
      )
    total_4m <- sum(data_4m[[value_column]], na.rm = TRUE)
    
    # Calculate 12-month window
    start_12m_ym <- block$start_12m_year * 12 + block$start_12m_month
    end_12m_ym <- block$end_12m_year * 12 + block$end_12m_month
    
    data_12m <- unit_data |>
      dplyr::filter(
        unit_data_ym >= start_12m_ym & unit_data_ym <= end_12m_ym
      )
    total_12m <- sum(data_12m[[value_column]], na.rm = TRUE)
    
    # Calculate seasonality percentage
    percent_seasonality <- ifelse(
      total_12m > 0, 
      (total_4m / total_12m) * 100, 
      0
    )
    is_seasonal <- as.numeric(percent_seasonality >= seasonality_threshold)
    
    # Create result row
    result_row <- data.frame(
      Block = i,
      DateRange = block$date_range,
      Total_4M = total_4m,
      Total_12M = total_12m,
      Percent_Seasonality = round(percent_seasonality, 2),
      Seasonal = is_seasonal,
      stringsAsFactors = FALSE
    )
    
    # Add administrative columns
    if (length(admin_columns) > 1) {
      admin_parts <- strsplit(admin_unit, " \\| ")[[1]]
      for (j in seq_along(admin_columns)) {
        if (j <= length(admin_parts)) {
          result_row[[admin_columns[j]]] <- admin_parts[j]
        } else {
          result_row[[admin_columns[j]]] <- NA
        }
      }
    } else {
      result_row[[admin_columns[1]]] <- admin_unit
    }
    
    detailed_results <- rbind(detailed_results, result_row)
  }
}

# Save results
writexl::write_xlsx(detailed_results, detailed_output)

# Display preview
knitr::kable(head(detailed_results), caption = "Preview of Detailed Block Results")
```

::: {.callout-note title="Output" icon="false"}
```{R}
#| message: false
#| echo: false
#| eval: true
# Initialize results dataframe
detailed_results <- data.frame()

# Get unique administrative groups
admin_groups <- unique(filtered_data$admin_group)

# Loop through each administrative unit
for (admin_unit in admin_groups) {
  # Filter data for this administrative unit
  unit_data <- filtered_data |>
    dplyr::filter(admin_group == admin_unit)
  
  # Loop through each time block
  for (i in 1:nrow(blocks)) {
    block <- blocks[i, ]
    
    # Create year-month values for comparison
    unit_data_ym <- unit_data[[year_column]] * 12 + unit_data$Month
    
    # Calculate 4-month window
    start_4m_ym <- block$start_4m_year * 12 + block$start_4m_month
    end_4m_ym <- block$end_4m_year * 12 + block$end_4m_month
    
    data_4m <- unit_data |>
      dplyr::filter(
        unit_data_ym >= start_4m_ym & unit_data_ym <= end_4m_ym
      )
    total_4m <- sum(data_4m[[value_column]], na.rm = TRUE)
    
    # Calculate 12-month window
    start_12m_ym <- block$start_12m_year * 12 + block$start_12m_month
    end_12m_ym <- block$end_12m_year * 12 + block$end_12m_month
    
    data_12m <- unit_data |>
      dplyr::filter(
        unit_data_ym >= start_12m_ym & unit_data_ym <= end_12m_ym
      )
    total_12m <- sum(data_12m[[value_column]], na.rm = TRUE)
    
    # Calculate seasonality percentage
    percent_seasonality <- ifelse(
      total_12m > 0, 
      (total_4m / total_12m) * 100, 
      0
    )
    is_seasonal <- as.numeric(percent_seasonality >= seasonality_threshold)
    
    # Create result row
    result_row <- data.frame(
      Block = i,
      DateRange = block$date_range,
      Total_4M = total_4m,
      Total_12M = total_12m,
      Percent_Seasonality = round(percent_seasonality, 2),
      Seasonal = is_seasonal,
      stringsAsFactors = FALSE
    )
    
    # Add administrative columns
    if (length(admin_columns) > 1) {
      admin_parts <- strsplit(admin_unit, " \\| ")[[1]]
      for (j in seq_along(admin_columns)) {
        if (j <= length(admin_parts)) {
          result_row[[admin_columns[j]]] <- admin_parts[j]
        } else {
          result_row[[admin_columns[j]]] <- NA
        }
      }
    } else {
      result_row[[admin_columns[1]]] <- admin_unit
    }
    
    detailed_results <- rbind(detailed_results, result_row)
  }
}

# Save results
writexl::write_xlsx(detailed_results, detailed_output)

# Display preview
knitr::kable(head(detailed_results), caption = "Preview of Detailed Block Results")
```

:::

**To adapt the code:**

- Do not change anything in the code 

### Step 6: Generate yearly summary

```{R}
#| message: false
#| echo: true
#| eval: false
#| code-fold: false
#| code-summary: "Show the code"
# Extract start year from DateRange
detailed_results$StartYear <- sapply(detailed_results$DateRange, function(x) {
  parts <- strsplit(x, "-")[[1]]
  first_part <- trimws(parts[1])
  as.numeric(substr(first_part, nchar(first_part) - 3, nchar(first_part)))
})

# Group by administrative columns and year
yearly_summary <- detailed_results |>
  dplyr::group_by(dplyr::across(dplyr::all_of(admin_columns)), StartYear) |>
  dplyr::summarise(
    Year = dplyr::first(StartYear),
    SeasonalCount = sum(Seasonal, na.rm = TRUE),
    total_blocks_in_year = 12,
    at_least_one_seasonal_block = as.numeric(SeasonalCount > 0),
    .groups = 'drop'
  ) |>
  dplyr::mutate(
    year_period = paste0(
      "(Jan ", Year, "-Apr ", Year, ", Dec ", Year, "-Mar ", Year + 1, ")"
    )
  ) |>
  dplyr::select(
    Year, 
    dplyr::all_of(admin_columns), 
    year_period, 
    total_blocks_in_year, 
    at_least_one_seasonal_block
  ) |>
  dplyr::arrange(Year, dplyr::across(dplyr::all_of(admin_columns)))

# Save results
writexl::write_xlsx(yearly_summary, yearly_output)

# Display results
knitr::kable(
  tail(yearly_summary, 10), 
  caption = "Yearly Seasonality Summary (Last 10 rows)"
)
```


::: {.callout-note title="Output" icon="false"}
```{R}
#| message: false
#| echo: false
#| eval: true
# Extract start year from DateRange
detailed_results$StartYear <- sapply(detailed_results$DateRange, function(x) {
  parts <- strsplit(x, "-")[[1]]
  first_part <- trimws(parts[1])
  as.numeric(substr(first_part, nchar(first_part) - 3, nchar(first_part)))
})

# Group by administrative columns and year
yearly_summary <- detailed_results |>
  dplyr::group_by(dplyr::across(dplyr::all_of(admin_columns)), StartYear) |>
  dplyr::summarise(
    Year = dplyr::first(StartYear),
    SeasonalCount = sum(Seasonal, na.rm = TRUE),
    total_blocks_in_year = 12,
    at_least_one_seasonal_block = as.numeric(SeasonalCount > 0),
    .groups = 'drop'
  ) |>
  dplyr::mutate(
    year_period = paste0(
      "(Jan ", Year, "-Apr ", Year, ", Dec ", Year, "-Mar ", Year + 1, ")"
    )
  ) |>
  dplyr::select(
    Year, 
    dplyr::all_of(admin_columns), 
    year_period, 
    total_blocks_in_year, 
    at_least_one_seasonal_block
  ) |>
  dplyr::arrange(Year, dplyr::across(dplyr::all_of(admin_columns)))

# Save results
writexl::write_xlsx(yearly_summary, yearly_output)

# Display results
knitr::kable(
  tail(yearly_summary, 10), 
  caption = "Yearly Seasonality Summary (Last 10 rows)"
)
```


:::

**To adapt the code:**

- Do not change anything in the code 

### Step 7: Location-level seasonality classification

```{R}
#| message: false
#| echo: true
#| eval: false
#| code-fold: false
#| code-summary: "Show the code"
# Group by all administrative columns and calculate seasonality classification
location_summary <- yearly_summary |>
  dplyr::group_by(dplyr::across(dplyr::all_of(admin_columns))) |>
  dplyr::summarise(
    SeasonalYears = sum(at_least_one_seasonal_block, na.rm = TRUE),
    TotalYears = dplyr::n(),
    .groups = 'drop'
  ) |>
  dplyr::mutate(
    Seasonality = ifelse(
      SeasonalYears == TotalYears, 
      "Seasonal", 
      "Not Seasonal"
    )
  ) |>
  dplyr::arrange(dplyr::across(dplyr::all_of(admin_columns)))

# Save results
writexl::write_xlsx(location_summary, location_output)

# Display results
knitr::kable(
  head(location_summary, 10), 
  caption = "Location Seasonality Classification"
)
```

::: {.callout-note title="Output" icon="false"}

```{R}
#| message: false
#| echo: false
#| eval: true
# Group by all administrative columns and calculate seasonality classification
location_summary <- yearly_summary |>
  dplyr::group_by(dplyr::across(dplyr::all_of(admin_columns))) |>
  dplyr::summarise(
    SeasonalYears = sum(at_least_one_seasonal_block, na.rm = TRUE),
    TotalYears = dplyr::n(),
    .groups = 'drop'
  ) |>
  dplyr::mutate(
    Seasonality = ifelse(
      SeasonalYears == TotalYears, 
      "Seasonal", 
      "Not Seasonal"
    )
  ) |>
  dplyr::arrange(dplyr::across(dplyr::all_of(admin_columns)))

# Save results
writexl::write_xlsx(location_summary, location_output)

# Display results
knitr::kable(
  head(location_summary, 10), 
  caption = "Location Seasonality Classification"
)
```

:::
