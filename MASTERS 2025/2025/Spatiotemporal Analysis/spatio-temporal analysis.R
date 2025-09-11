# 📦 Install and load necessary packages
library(readxl)
library(tidyverse)
library(spacetime)
library(sp)
library(gstat)
library(mgcv)

# 📂 Combine all Excel files into a single data frame
file_paths <- list.files(path = "C:\\Users\\AdonsMB\\OneDrive - University of the Free State\\Desktop\\Masters\\MASTERS 2025\\2025\\Data\\Complete Imputed Data", pattern = "*.xlsx", full.names = TRUE)

# This new loop will handle the monthly data structure
all_data <- map_dfr(file_paths, function(path) {
  station_name <- tools::file_path_sans_ext(basename(path))
  
  # Read the Precipitation sheet and clean month names
  precip_data <- read_excel(path, sheet = "Precipitation") %>%
    pivot_longer(
      cols = -Year,
      names_to = "Month_Name",
      values_to = "Precipitation"
    ) %>%
    mutate(Station = station_name)
  
  # Read the Maximum Temperature sheet and clean month names
  temp_data <- read_excel(path, sheet = "Maximum") %>%
    pivot_longer(
      cols = -Year,
      names_to = "Month_Name",
      values_to = "Max_Temperature"
    ) %>%
    mutate(Station = station_name)
  
  # Join precipitation and temperature data by Year, Month_Name, and Station
  full_data <- left_join(precip_data, temp_data, by = c("Year", "Month_Name", "Station"))
  
  # Add month number, making the matching more robust
  full_data <- full_data %>%
    mutate(Month = match(
      tolower(str_trim(Month_Name)), # Make month names lowercase and remove whitespace
      tolower(month.abb)            # Match against lowercase abbreviations
    ))
  
  return(full_data)
})

# 🔗 Merge with coordinates and format data
stations_info <- read_excel("C:\\Users\\AdonsMB\\OneDrive - University of the Free State\\Desktop\\Masters\\MASTERS 2025\\2025\\FS Coordinates.xlsx")

# Create a temporary data frame
tmp_data <- all_data %>%
  left_join(stations_info, by = "Station") %>%
  select(-Month_Name)

# Manually calculate the scaling parameters (mean and standard deviation)
lon_center <- mean(tmp_data$Longitude, na.rm = TRUE)
lon_scale <- sd(tmp_data$Longitude, na.rm = TRUE)
lat_center <- mean(tmp_data$Latitude, na.rm = TRUE)
lat_scale <- sd(tmp_data$Latitude, na.rm = TRUE)

# Apply the scaling to the temporary data frame
tmp_data <- tmp_data %>%
  mutate(
    Latitude_sc = (Latitude - lat_center) / lat_scale,
    Longitude_sc = (Longitude - lon_center) / lon_scale
  )

# Create the final data object for modeling
final_data <- tmp_data

# Display a sample of the final data frame
head(final_data)

# 🌍 Convert your data frame to a spatial points data frame
coordinates(final_data) <- ~Longitude + Latitude

# --- END OF MODIFIED CODE ---

#############################################################################################################
#############################################################################################################

# 🔎 Create a variogram for Precipitation
precip_vgm <- variogram(Precipitation ~ 1, data = final_data)

# 📈 Plot the variogram
plot(precip_vgm, main = "Variogram for Precipitation",col='blue',pch=16)

# 🔎 Create a variogram for Temperature
temp_vgm <- variogram(`Max_Temperature` ~ 1, data = final_data)

# 📈 Plot the variogram
plot(temp_vgm, main = "Variogram for Temperature",col='blue',pch=16)

###############################################################################################################
###############################################################################################################

# ☔️ Fit a spatio-temporal model for Precipitation
gam_precip <- gam(Precipitation ~ s(Year) + s(Longitude_sc, Latitude_sc,k=7) + 
                    ti(Longitude_sc, Latitude_sc, Year), 
                  data = final_data, 
                  family = quasipoisson)

# 🌡️ Fit a spatio-temporal model for Maximum Temperature
gam_temp <- gam(Max_Temperature ~ s(Year) + s(Longitude_sc, Latitude_sc,k=7) + 
                  ti(Longitude_sc, Latitude_sc, Year), 
                data = final_data, 
                family = gaussian)

# 📝 Print the summaries to get the results
summary(gam_precip)
summary(gam_temp)

gam_precip_spatial = gam_precip
gam_temp_spatial = gam_temp

###########################################################################################################
###########################################################################################################

# 🌐 Create a spatial grid for prediction
# Get the spatial extent (bounding box) of your data
grid_extent <- expand.grid(
  Longitude = seq(min(final_data$Longitude), max(final_data$Longitude), length.out = 100),
  Latitude = seq(min(final_data$Latitude), max(final_data$Latitude), length.out = 100)
)

# Convert the grid to a spatial points data frame
coordinates(grid_extent) <- ~Longitude + Latitude

# Now, create the 'grid_points' object
grid_points <- as.data.frame(grid_extent)

# Manually scale the new grid coordinates using the stored parameters
grid_points$Longitude_sc <- (grid_points$Longitude - lon_center) / lon_scale
grid_points$Latitude_sc <- (grid_points$Latitude - lat_center) / lat_scale

# Add a Year column. For a static prediction map, you can use a single year, e.g., the last year of your data
grid_points$Year <- max(final_data$Year)

# 🗺️ Predict precipitation and temperature values across the spatial grid
grid_points$precip_pred <- predict(gam_precip_spatial, newdata = as.data.frame(grid_points), type = "response")
grid_points$temp_pred <- predict(gam_temp_spatial, newdata = as.data.frame(grid_points), type = "response")

# 📊 Convert the grid back to a data frame for plotting with ggplot2
grid_df <- as.data.frame(grid_points)
# Convert the final_data back to a data frame for plotting
final_data_df <- as.data.frame(final_data)

# 📈 Visualize the spatial predictions for Precipitation with station names
ggplot(grid_df, aes(x = Longitude, y = Latitude, fill = precip_pred)) +
  geom_raster() +
  geom_point(data = final_data_df, aes(x = Longitude, y = Latitude), color = "yellow", shape = 20, size = 3, inherit.aes = FALSE) +
  # Add text labels for station names
  geom_text(data = final_data_df, aes(x = Longitude, y = Latitude, label = Station),
            vjust = -1, hjust = 0.5, size = 3, color = "yellow", inherit.aes = FALSE) + # Adjust vjust/hjust for label position
  labs(title = "Predicted Spatial Surface for Precipitation",
       subtitle = "Station locations with names",
       x = "Longitude", y = "Latitude") +
  theme_minimal()

# 📈 Visualize the spatial predictions for Maximum Temperature with station names
ggplot(grid_df, aes(x = Longitude, y = Latitude, fill = temp_pred)) +
  geom_raster() +
  geom_point(data = final_data_df, aes(x = Longitude, y = Latitude), color = "yellow", shape = 20, size = 3, inherit.aes = FALSE) +
  # Add text labels for station names
  geom_text(data = final_data_df, aes(x = Longitude, y = Latitude, label = Station),
            vjust = -1, hjust = 0.5, size = 3, color = "yellow", inherit.aes = FALSE) + # Adjust vjust/hjust for label position
  labs(title = "Predicted Spatial Surface for Maximum Temperature",
       subtitle = "Station locations with names",
       x = "Longitude", y = "Latitude") +
  theme_minimal()