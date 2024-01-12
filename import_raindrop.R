pacman::p_load(tidyverse,
               readxl,
               tidyxl,
               glue,
               janitor)

# The ONS data contains hyperlinks which must be disaggregated into text nd url using a function in the data/ONS_hyperlinks.xlsm spreadsheet first

ons_path <- "data/ONS_hyperlinks.xlsx"


a_team_tbl <- read_xlsx("data/Analysis Team Data Library.xlsx",
                        sheet = "Library ", range = "B2:k144")

reports <- read_xlsx("data/Analysis Team Data Library.xlsx",
                     sheet = "Reports", range = "A2:J44")

tags_tbl <- read_xlsx("data/Analysis Team Data Library.xlsx",
                      sheet = "Dropdown", range = "A1:A17", col_names = "tag")

# ONS ----
ons_sheet_list <- excel_sheets(ons_path)

ons_tables_list <- map2(.x = ons_path, 
                        .y = ons_sheet_list,
                        .f = ~read_xlsx(.x, .y, range = "A2:AN44")) %>% 
  set_names(ons_sheet_list)


clean_wrangle_ons_tbl <- function(ons_tbl, sheet_name){
# another wildly inconsistent spreadsheet
df1 <- ons_tbl %>%
  select(!starts_with("..")) %>%
  clean_names() %>% 
  relocate(url, .after = link) %>% 
  relocate(starts_with("date"), .before = link) %>% 
  relocate(frequency, .before = link) %>% 
  remove_empty("rows") %>% 
  filter(!is.na(url))

if("geography" %in% names(df1)){
  df1 <- df1 %>% relocate(geography, .before = link)
} else {
  df1 <- df1 %>% mutate(geography = "") %>% 
    relocate(geography, .before = link)
}

tagscol_start <- which(names(df1) %>% tolower() == "url") + 1
df <-  df1 %>% 
  select(!where(~all(as.character(.x) %in% c("No","N/A"), na.rm = TRUE))) %>%
  mutate(across(tagscol_start:last_col(),
                ~ if_else(as.character(.x) == "Yes",
                          cur_column(),
                          "")
                )
         ) %>%
  unite(col = tags, tagscol_start:last_col(), remove = TRUE, sep = ",") %>%
  mutate(tags = str_replace_all(tags, "NA|,+", ",") %>%
           str_replace("^,", "") %>% 
           str_replace(",+$", ""),
         topic = str_replace_all(sheet_name, "(?<=[a-z])(?=[A-Z])", " "))
return(df)
}


make_ons_output_tbl <- function(ons_clean_tbl){

  ons_clean_tbl %>% 
  transmute(url,
            folder = glue("Bookmarks bar/work/WECA/ONS/{topic}"),
            title = glue("{measure}: {link}"),
            note = glue("{frequency}\n{geography}\n{source}"),
            tags = glue("{tags},{theme}"),
            created = date_of_last_update %>% as.character(),
            )
}

# take the named list, clean the data and use the list name in the folder path
cleaned_dfs <- imap(ons_tables_list, clean_wrangle_ons_tbl)

# make the import file
output <- map(cleaned_dfs, make_ons_output_tbl) %>% 
  bind_rows()

output %>% view()

write_csv(output, "data/ons_out.csv")  

# ANALYSIS TEAM ----
import_tbl <- tribble(~url, ~folder, ~title, ~note, ~tags, ~created)

a_team_output <- a_team_tbl %>% 
  clean_names %>% 
  filter(str_starts(source, "http")) %>% 
  transmute(url = source,
            folder = glue("Bookmarks bar/work/WECA/A_team_data/{topic}"),
            title = glue("{data_on}: {name_of_data}"),
            tags = glue("{data_on},{geographies}"),
            note = notes_used_in,
            created = Sys.Date()
            )

write_csv(a_team_output, file = "data/a_team_output.csv", na = "")

