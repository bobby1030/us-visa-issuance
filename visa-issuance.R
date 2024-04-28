library(tidyverse)
library(readxl)

filepath <- "./FYs97-23_NIVDetailTable.xlsx"

sheetnames <- excel_sheets(filepath)
names(sheetnames) <- 1997:2023

visa_dfs <- imap(sheetnames, \(sheetname, year) {
    df <- read_xlsx(filepath, sheet = sheetname)
    df <- df %>%
        rename("nationality" = 1) %>% # rename the first column to "nationality"
        mutate(year = as.integer(year)) %>%
        relocate(year) %>%
        rename_with(\(x) str_remove(x, "-")) %>% # remove hyphens from visa types
        mutate(across(!nationality, as.numeric))
    return(df)
})

visa <- bind_rows(visa_dfs)

visa_f1 <- visa %>% 
    mutate(
        country_code = case_when(
            str_detect(nationality, "China - mainland") ~ "China",
            str_detect(nationality, "India") ~ "India",
            str_detect(nationality, "Taiwan") ~ "Taiwan",
            str_detect(nationality, "Japan") ~ "Japan",
            str_detect(nationality, "Korea, South") ~ "South Korea",
            .default = NA_character_
        )
    ) %>%
    filter(!is.na(country_code))

plt_visa_f1 <- ggplot(visa_f1, aes(x = year, y = F1, color = country_code)) +
    geom_line() +
    scale_x_continuous(breaks = seq(1998, 2023, 2)) +
    scale_y_continuous(labels = scales::label_number_auto()) +
    labs(
        title = "F1 Visa Issuance by Nationality",
        caption = "Data source: U.S. Department of State",
        x = "Fiscal Year",
        y = "Number of Visas Issued",
        color = "Nationality"
    )

ggsave("f1-visa-issuance.png", plt_visa_f1, width = 8, height = 5, dpi = 300)
