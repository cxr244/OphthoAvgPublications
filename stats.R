# Ophtho Avg Pub Project

library(dplyr) # For Data frame manipulation 
library("readxl") # Read Excel docs
library(gtsummary) # Create statistics tables 
library(officer) # Manipulate word docs
library(flextable) # Convert tables to word doc format 


my_data <- read_excel("C:/Chandru/CWRU/Research/CCF/OphthoAvgPubsPaper/StatsTest.xlsx", sheet="Sheet1")

# Make sure all applicants are accounted for, some are missing with 0 and not recorded in sus authors? 

Table1 <- my_data %>% select(NumberPubs, OphthoPubs, 
        FirstAuthorPubs, FirstAuthorOphthoPubs,
        HighestIF, AverageIF, Gender, TopMedSchool, TopDoximity) %>% 
        tbl_summary(missing= "no",
        statistic = list(all_continuous() ~ "{mean} \u00b1 {sd}"),
        label = list(NumberPubs ~ "Number of Publications", 
        OphthoPubs ~ "Ophthalmology Publications",
        FirstAuthorPubs ~ "First Author Publications",
        FirstAuthorOphthoPubs ~ "First Author Ophthalmology Publications",
        HighestIF ~ "Highest Impact Factor of Publications", 
        AverageIF ~ "Average Impact Factor of Publcations", 
        TopMedSchool ~ "Top 40 NIH Funded Medical School", 
        TopDoximity ~ "Top 20 Doximity Ranked Program"), 
        type = list(c(NumberPubs, OphthoPubs, 
        FirstAuthorPubs, FirstAuthorOphthoPubs,
        HighestIF, AverageIF) ~ "continuous"),
        digits = all_continuous() ~ 1) %>%
        modify_spanning_header(c("stat_1", "stat_2") ~ "**Treatment Received**")%>%
        add_n() %>% bold_labels() 

Table2 <- my_data %>% select(NumberPubs, OphthoPubs, 
        FirstAuthorPubs, FirstAuthorOphthoPubs, 
        HighestIF, AverageIF, Gender, TopMedSchool, TopDoximity) %>% 
        tbl_summary(by=Gender, missing= "no",
        statistic = list(all_continuous() ~ "{mean} \u00b1 {sd}"),
        label = list(NumberPubs ~ "Number of Publications", 
        OphthoPubs ~ "Ophthalmology Publications",
        FirstAuthorPubs ~ "First Author Publications",
        FirstAuthorOphthoPubs ~ "First Author Ophthalmology Publications",
        HighestIF ~ "Highest Impact Factor of Publications", 
        AverageIF ~ "Average Impact Factor of Publcations", 
        TopMedSchool ~ "Top 40 NIH Funded Medical School", 
        TopDoximity ~ "Top 20 Doximity Ranked Program"), 
        type = list(c(NumberPubs, OphthoPubs, 
        FirstAuthorPubs, FirstAuthorOphthoPubs, 
        HighestIF, AverageIF) ~ "continuous"),
        digits = all_continuous() ~ 1) %>%
        add_n() %>% add_p(test = list(all_continuous() ~ "t.test")) %>% bold_labels() %>% bold_p() 
Table3 <- my_data %>% select(NumberPubs, OphthoPubs, 
        FirstAuthorPubs, FirstAuthorOphthoPubs, 
        HighestIF, AverageIF, Gender, TopMedSchool, TopDoximity) %>% 
        tbl_summary(by=TopMedSchool, missing= "no",
        statistic = list(all_continuous() ~ "{mean} \u00b1 {sd}"),
        label = list(NumberPubs ~ "Number of Publications", 
        OphthoPubs ~ "Ophthalmology Publications",
        FirstAuthorPubs ~ "First Author Publications",
        FirstAuthorOphthoPubs ~ "First Author Ophthalmology Publications",
        HighestIF ~ "Highest Impact Factor of Publications", 
        AverageIF ~ "Average Impact Factor of Publcations", 
        TopMedSchool ~ "Top 40 NIH Funded Medical School", 
        TopDoximity ~ "Top 20 Doximity Ranked Program"), 
        type = list(c(NumberPubs, OphthoPubs, 
        FirstAuthorPubs, FirstAuthorOphthoPubs, 
        HighestIF, AverageIF) ~ "continuous"),
        digits = all_continuous() ~ 1) %>%
        add_n() %>% add_p(test = list(all_continuous() ~ "t.test")) %>% 
        bold_labels() %>% bold_p() 
Table4 <- my_data %>% select(NumberPubs, OphthoPubs, 
        FirstAuthorPubs, FirstAuthorOphthoPubs, 
        HighestIF, AverageIF, Gender, TopMedSchool, TopDoximity) %>% 
        tbl_summary(by=TopDoximity, missing= "no",
        statistic = list(all_continuous() ~ "{mean} \u00b1 {sd}"),
        label = list(NumberPubs ~ "Number of Publications", 
        OphthoPubs ~ "Ophthalmology Publications",
        FirstAuthorPubs ~ "First Author Publications",
        FirstAuthorOphthoPubs ~ "First Author Ophthalmology Publications",
        HighestIF ~ "Highest Impact Factor of Publications", 
        AverageIF ~ "Average Impact Factor of Publcations", 
        TopMedSchool ~ "Top 40 NIH Funded Medical School", 
        TopDoximity ~ "Top 20 Doximity Ranked Program"), 
        type = list(c(NumberPubs, OphthoPubs, 
        FirstAuthorPubs, FirstAuthorOphthoPubs, 
        HighestIF, AverageIF) ~ "continuous"),
        digits = all_continuous() ~ 1) %>%
        add_n() %>% add_p(test = list(all_continuous() ~ "t.test")) %>% bold_labels() %>% bold_p() 

toptwenty <- glm(TopDoximity ~ NumberPubs + OphthoPubs + 
        FirstAuthorPubs + Gender + FirstAuthorOphthoPubs + 
        HighestIF + AverageIF + TopMedSchool, data = my_data,
						family = binomial())
Table5 <- toptwenty %>% tbl_regression(exponentiate = TRUE, 
		label = list(NumberPubs ~ "Number of Publications", 
        OphthoPubs ~ "Ophthalmology Publications",
        FirstAuthorPubs ~ "First Author Publications",
        FirstAuthorOphthoPubs ~ "First Author Ophthalmology Publications",
        HighestIF ~ "Highest Impact Factor of Publications", 
        AverageIF ~ "Average Impact Factor of Publcations", 
        TopMedSchool ~ "Top 40 NIH Funded Medical School"),
        Gender ~ "Gender (Reference = Female)",
        digits = all_continuous() ~ 1) %>% 
		bold_labels() %>% 
		bold_p(t = 0.05) %>% modify_header(estimate = "Odds Ratio")

  doc <- read_docx() 
  doc <- body_add_par(doc, value="Table 1 - Demographics", style="heading 1")
  doc <- body_add_flextable(doc,value= Table1 %>% gtsummary::as_flextable()) %>% body_add_break()

  doc <- body_add_par(doc, value="Table 2 - Demographics by Gender", style="heading 1")
  doc <- body_add_flextable(doc,value= Table2 %>% gtsummary::as_flextable()) %>% body_add_break()

  doc <- body_add_par(doc, value="Table 3 - Demographics by Top 40 NIH Med School", style="heading 1")
  doc <- body_add_flextable(doc,value= Table3 %>% gtsummary::as_flextable()) %>% body_add_break()

  doc <- body_add_par(doc, value="Table 4 - Demographics by Top 20 Doximity Ranked Program", style="heading 1")
  doc <- body_add_flextable(doc,value= Table4 %>% gtsummary::as_flextable()) %>% body_add_break()

  doc <- body_add_par(doc, value="Table 5 - Regression Model for association with matching at Top 20 Doximity Program", style="heading 1")
  doc <- body_add_flextable(doc,value= Table5 %>% gtsummary::as_flextable()) %>% body_add_break()


  print(doc, target="C:/Chandru/CWRU/Research/CCF/OphthoAvgPubsPaper/Tables.docx")
  




