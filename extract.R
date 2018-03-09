library(tidyverse)
library(googledrive)
library(tidyxl)
library(unpivotr)
library(crayon)
library(here)

path <- here("taxonomy.xlsx")

drive_download(file = as_id("1zdYAEgH93QzqktNRytK0nABMDQ9yBJ4BmKvlvUTU9JA"),
               path = path,
               overwrite = TRUE)

sheet_name <- "All themes combined"

formats <- xlsx_formats(path)

font_colours <- formats$local$font$color$rgb
fill_colours <- formats$local$fill$patternFill$bgColor$rgb

cells <-
  xlsx_cells(path, sheet_name) %>%
  filter(!is_blank) %>%
  select(row, col, address, character, character_formatted, local_format_id) %>%
  mutate(is_multiformatted = map_int(character_formatted, nrow) != 1,
         font_colour = font_colours[local_format_id],
         fill_colour = fill_colours[local_format_id],
         character = str_replace(character, "\\n", ""))

data_cells <- filter(cells, row >= 2, col <= 4)

font_colour_legend <-
  cells %>%
  filter(col == 5, row %in% 2:7) %>%
  select(group = character, local_format_id) %>%
  mutate(font_colour = font_colours[local_format_id]) %>%
  select(-local_format_id)

fill_colour_legend <-
  cells %>%
  filter(col == 5, row %in% 9:10) %>%
  select(status = character, local_format_id) %>%
  mutate(fill_colour = fill_colours[local_format_id]) %>%
  select(-local_format_id)

# Check for cells with colours that aren't in the legend

data_cells %>%
  group_by(fill_colour) %>%
  summarise(address = first(address)) %>%
  anti_join(fill_colour_legend, by = "fill_colour")
# FFB7B7B7 grey = "Done"
# FFD9D9D9 grey = "Done"
# FFFFFFFF white = "Not done" but ignore because "Not done" isn't in the legend.

# Add the missing fills to the legend
fill_colour_legend <-
  bind_rows(fill_colour_legend,
            tibble(fill_colour = c("FFB7B7B7", "FFD9D9D9"),
                   status = "Done"))

unicoloured_strings <-
  data_cells %>%
  filter(!is_multiformatted) %>%
  select(row, col, address, character, font_colour, fill_colour)

multicoloured_strings <-
  data_cells %>%
  filter(is_multiformatted) %>%
  select(row, col, address, character_formatted, font_colour, fill_colour) %>%
  unnest() %>%
  select(row, col, address, character, color_rgb, font_colour, fill_colour) %>%
  mutate(character = str_trim(str_replace(character, "AND", "")),
         font_colour = if_else(is.na(color_rgb), font_colour, color_rgb)) %>%
  select(-color_rgb) %>%
  filter(character != "")

# Cheeck that all whole-cell font colours are in the legend.  Whole-cell font
# colours are the default colour of substrings of a multicoloured cell.
data_cells %>%
  anti_join(font_colour_legend, by = "font_colour") %>%
  select(address, font_colour)
# FF38761D green = "Policy area"

# Font colours in multicoloured cells that aren't in the legend
anti_join(multicoloured_strings, font_colour_legend, by = "font_colour")
# FFE69138 orange = "Policy"
# FFF1C232 orange = "Policy"

# Add the missing fonts to the legend
font_colour_legend <-
  bind_rows(font_colour_legend,
            tibble(font_colour = c("FF38761D", "FFE69138", "FFF1C232"),
                   group = c("Policy area", "Policy", "Policy")))

all_strings <- bind_rows(unicoloured_strings, multicoloured_strings)

tags <-
  all_strings %>%
  left_join(font_colour_legend, by = "font_colour") %>%
  left_join(fill_colour_legend, by = "fill_colour") %>%
  transmute(level = col,
            row,
            col,
            tag = character,
            group,
            status = if_else(is.na(status), "Not done", status),
            address)

# Mark the parent tag of each tag

level1 <- filter(tags, level == 1)
level2 <- filter(tags, level == 2)
level3 <- filter(tags, level == 3)
level4 <- filter(tags, level == 4)

level4parented <- WNW(level4, select(level3, row, col, parent = address))
level3parented <- WNW(level3, select(level2, row, col, parent = address))
level2parented <- WNW(level2, select(level1, row, col, parent = address))

parented <-
  bind_rows(level1, level2parented, level3parented, level4parented) %>%
  select(address, level, row, tag, group, status, parent) %>%
  arrange(row, level)

# Write to disk
write_tsv(parented, here("tags.tsv"))

# Basic sense checks

# No NAs
anyNA(select(parented, -parent))

# Some duplicates
parented %>%
  count(tag, sort = TRUE) %>%
  filter(n > 1) %>%
  left_join(tags, by = "tag") %>%
  select(-n) %>%
  print(n = Inf)
