# Ophtho Avg Pub Project

library(dplyr) # For Data frame manipulation 
library("readxl") # Read Excel docs
library(gtsummary) # Create statistics tables 
library(officer) # Manipulate word docs
library(flextable) # Convert tables to word doc format 


my_data <- read_excel("C:/Chandru/CWRU/Research/CCF/OphthoAvgPubsPaper/StatsTest.xlsx", sheet="Sheet1")

# Make sure all applicants are accounted for, some are missing with 0 and not recorded in sus authors? 

my_data <- my_data %>% mutate(
         PubsInProgress = NumberPubs - PubBeforeSubmission,
         OphthoInProgress = OphthoPubs - OphthoSubmission,
         FirstAuthorInProgress = FirstAuthorPubs - FirstAuthorSubmission,
         FirstAuthorOphthoInProgress = FirstAuthorOphthoPubs - FirstAuthorOphthoSubmission)



# New Huge Table 2 
Table1new <- my_data %>% select(Gender, TopMedSchool, TopDoximity, PubBeforeSubmission, OphthoSubmission,
        FirstAuthorSubmission, FirstAuthorOphthoSubmission, HighestIFSubmission, AverageIFSubmission, 
        PubsInProgress, OphthoInProgress, FirstAuthorInProgress, FirstAuthorOphthoInProgress, 
        HighestIF, AverageIF) %>% 
        tbl_summary(missing= "no",
        statistic = list(all_continuous() ~ "{mean} \u00b1 {sd}"),
        label = list(
        Gender ~ "Gender",
        TopMedSchool ~ "Top 40 US News Medical School: Research", 
        TopDoximity ~ "Top 20 Doximity Ranked Program",
        PubBeforeSubmission ~ "Number of Publications", 
        OphthoSubmission ~ "Ophthalmology Publications",
        FirstAuthorSubmission ~ "First Author Publications",
        FirstAuthorOphthoSubmission ~ "First Author Ophthalmology Publications",
        HighestIFSubmission ~ "Highest Impact Factor of Publications",
        AverageIFSubmission ~ "Average Impact Factor of Publications",         
        PubsInProgress ~ "Number Of Publications",        
        OphthoInProgress ~ "Ophthalmology Publications",        
        FirstAuthorInProgress ~ "First Author Publications",        
        FirstAuthorOphthoInProgress ~ "First Author Ophthalmology Publications",
        HighestIF ~ "Highest Impact Factor of Publications",         
        AverageIF ~ "Average Impact Factor of Publications"), 
        type = list(c(PubBeforeSubmission, PubsInProgress, OphthoSubmission,
        OphthoInProgress, FirstAuthorSubmission, FirstAuthorInProgress, 
        FirstAuthorOphthoSubmission,
        FirstAuthorOphthoInProgress, HighestIF, HighestIFSubmission, 
        AverageIF, AverageIFSubmission) ~ "continuous"),
        digits = all_continuous() ~ 1) %>%
        bold_labels() 
Table2new <- my_data %>% select(Gender, TopMedSchool, TopDoximity, PubBeforeSubmission, OphthoSubmission,
        FirstAuthorSubmission, FirstAuthorOphthoSubmission, HighestIFSubmission, AverageIFSubmission, 
        PubsInProgress, OphthoInProgress, FirstAuthorInProgress, FirstAuthorOphthoInProgress, 
        HighestIF, AverageIF) %>% 
        tbl_summary(by=Gender, missing= "no",
        statistic = list(all_continuous() ~ "{mean} \u00b1 {sd}"),
        label = list(
        Gender ~ "Gender",
        TopMedSchool ~ "Top 40 US News Medical School: Research", 
        TopDoximity ~ "Top 20 Doximity Ranked Program",
        PubBeforeSubmission ~ "Number of Publications", 
        OphthoSubmission ~ "Ophthalmology Publications",
        FirstAuthorSubmission ~ "First Author Publications",
        FirstAuthorOphthoSubmission ~ "First Author Ophthalmology Publications",
        HighestIFSubmission ~ "Highest Impact Factor of Publications",
        AverageIFSubmission ~ "Average Impact Factor of Publications",         
        PubsInProgress ~ "Number Of Publications",        
        OphthoInProgress ~ "Ophthalmology Publications",        
        FirstAuthorInProgress ~ "First Author Publications",        
        FirstAuthorOphthoInProgress ~ "First Author Ophthalmology Publications",
        HighestIF ~ "Highest Impact Factor of Publications",         
        AverageIF ~ "Average Impact Factor of Publications"), 
        type = list(c(PubBeforeSubmission, PubsInProgress, OphthoSubmission,
        OphthoInProgress, FirstAuthorSubmission, FirstAuthorInProgress, 
        FirstAuthorOphthoSubmission,
        FirstAuthorOphthoInProgress, HighestIF, HighestIFSubmission, 
        AverageIF, AverageIFSubmission) ~ "continuous"),
        digits = all_continuous() ~ 1) %>%
        add_p(test = list(all_continuous() ~ "t.test")) %>% bold_labels() %>% bold_p()
Table3new <- my_data %>% select(Gender, TopMedSchool, TopDoximity, PubBeforeSubmission, OphthoSubmission,
        FirstAuthorSubmission, FirstAuthorOphthoSubmission, HighestIFSubmission, AverageIFSubmission, 
        PubsInProgress, OphthoInProgress, FirstAuthorInProgress, FirstAuthorOphthoInProgress, 
        HighestIF, AverageIF) %>% 
        tbl_summary(by=TopMedSchool, missing= "no",
        statistic = list(all_continuous() ~ "{mean} \u00b1 {sd}"),
        label = list(
        Gender ~ "Gender",
        TopMedSchool ~ "Top 40 US News Medical School: Research", 
        TopDoximity ~ "Top 20 Doximity Ranked Program",
        PubBeforeSubmission ~ "Number of Publications", 
        OphthoSubmission ~ "Ophthalmology Publications",
        FirstAuthorSubmission ~ "First Author Publications",
        FirstAuthorOphthoSubmission ~ "First Author Ophthalmology Publications",
        HighestIFSubmission ~ "Highest Impact Factor of Publications",
        AverageIFSubmission ~ "Average Impact Factor of Publications",         
        PubsInProgress ~ "Number Of Publications",        
        OphthoInProgress ~ "Ophthalmology Publications",        
        FirstAuthorInProgress ~ "First Author Publications",        
        FirstAuthorOphthoInProgress ~ "First Author Ophthalmology Publications",
        HighestIF ~ "Highest Impact Factor of Publications",         
        AverageIF ~ "Average Impact Factor of Publications"), 
        type = list(c(PubBeforeSubmission, PubsInProgress, OphthoSubmission,
        OphthoInProgress, FirstAuthorSubmission, FirstAuthorInProgress, 
        FirstAuthorOphthoSubmission,
        FirstAuthorOphthoInProgress, HighestIF, HighestIFSubmission, 
        AverageIF, AverageIFSubmission) ~ "continuous"),
        digits = all_continuous() ~ 1) %>%
        add_p(test = list(all_continuous() ~ "t.test")) %>% bold_labels() %>% bold_p() 
Table4new <- my_data %>% select(Gender, TopMedSchool, TopDoximity, PubBeforeSubmission, OphthoSubmission,
        FirstAuthorSubmission, FirstAuthorOphthoSubmission, HighestIFSubmission, AverageIFSubmission, 
        PubsInProgress, OphthoInProgress, FirstAuthorInProgress, FirstAuthorOphthoInProgress, 
        HighestIF, AverageIF) %>% 
        tbl_summary(by=TopDoximity, missing= "no",
        statistic = list(all_continuous() ~ "{mean} \u00b1 {sd}"),
        label = list(
        Gender ~ "Gender",
        TopMedSchool ~ "Top 40 US News Medical School: Research", 
        TopDoximity ~ "Top 20 Doximity Ranked Program",
        PubBeforeSubmission ~ "Number of Publications", 
        OphthoSubmission ~ "Ophthalmology Publications",
        FirstAuthorSubmission ~ "First Author Publications",
        FirstAuthorOphthoSubmission ~ "First Author Ophthalmology Publications",
        HighestIFSubmission ~ "Highest Impact Factor of Publications",
        AverageIFSubmission ~ "Average Impact Factor of Publications",         
        PubsInProgress ~ "Number Of Publications",        
        OphthoInProgress ~ "Ophthalmology Publications",        
        FirstAuthorInProgress ~ "First Author Publications",        
        FirstAuthorOphthoInProgress ~ "First Author Ophthalmology Publications",
        HighestIF ~ "Highest Impact Factor of Publications",         
        AverageIF ~ "Average Impact Factor of Publications"), 
        type = list(c(PubBeforeSubmission, PubsInProgress, OphthoSubmission,
        OphthoInProgress, FirstAuthorSubmission, FirstAuthorInProgress, 
        FirstAuthorOphthoSubmission,
        FirstAuthorOphthoInProgress, HighestIF, HighestIFSubmission, 
        AverageIF, AverageIFSubmission) ~ "continuous"),
        digits = all_continuous() ~ 1) %>%
        add_p(test = list(all_continuous() ~ "t.test")) %>% bold_labels() %>% bold_p() 

CombinedTable12 <- tbl_merge (
                tbls = list(Table1new, Table2new), 
                tab_spanner = c("**Overall**", "**Gender**")) %>% 
                gtsummary::as_flextable()
CombinedTable34 <- tbl_merge (
                tbls = list(Table3new, Table4new), 
                tab_spanner = c("**Top 40 Medical School**", "**Top 20 Doximity Program**")) %>% 
                gtsummary::as_flextable()






# Table 1 Alternate cleaner version 
Table1PreSubmission <- my_data %>% select(PubBeforeSubmission, OphthoSubmission,
        FirstAuthorSubmission, FirstAuthorOphthoSubmission,
        HighestIFSubmission, 
        AverageIFSubmission) %>% 
        tbl_summary(missing= "no",
        statistic = list(all_continuous() ~ "{mean} \u00b1 {sd}"),
        label = list(
        PubBeforeSubmission ~ "Number of Publications", 
        OphthoSubmission ~ "Ophthalmology Publications",
        FirstAuthorSubmission ~ "First Author Publications",
        FirstAuthorOphthoSubmission ~ "First Author Ophthalmology Publications",
        HighestIFSubmission ~ "Highest Impact Factor of Publications", 
        AverageIFSubmission ~ "Average Impact Factor of Publications"), 
        type = list(c(PubBeforeSubmission, OphthoSubmission,
        FirstAuthorSubmission, 
        FirstAuthorOphthoSubmission,
        HighestIFSubmission, 
        AverageIFSubmission) ~ "continuous"),
        digits = all_continuous() ~ 1) %>%
        bold_labels() 

Table1PostSubmission <- my_data %>% select(PubsInProgress,
        OphthoInProgress, FirstAuthorInProgress, 
        FirstAuthorOphthoInProgress, HighestIF, 
        AverageIF) %>% 
        tbl_summary(missing= "no",
        statistic = list(all_continuous() ~ "{mean} \u00b1 {sd}"),
        label = list(
        PubsInProgress ~ "Number Of Publications",
        OphthoInProgress ~ "Ophthalmology Publications",
        FirstAuthorInProgress ~ "First Author Publications",
        FirstAuthorOphthoInProgress ~ "First Author Ophthalmology Publications",
        HighestIF ~ "Highest Impact Factor of Publications",
        AverageIF ~ "Average Impact Factor of Publications"), 
        type = list(c(PubsInProgress,
        OphthoInProgress, 
        FirstAuthorInProgress, 
        FirstAuthorOphthoInProgress, 
        HighestIF, 
        AverageIF) ~ "continuous"),
        digits = all_continuous() ~ 1) %>%
        bold_labels() 
CombinedTable1 <- tbl_merge (
                tbls = list(Table1PreSubmission, Table1PostSubmission), 
                tab_spanner = c("**Before Submission**", "**Research In Progress**")) %>% 
                gtsummary::as_flextable()


data1 <- my_data %>% select(PubBeforeSubmission, OphthoSubmission, FirstAuthorSubmission,
                            FirstAuthorOphthoSubmission)
pc1 <- prcomp(data1, scale. = FALSE)
pc1$rotation
summary(pc1)

data2 <- my_data %>% select(PubsInProgress, OphthoInProgress, 
                            FirstAuthorInProgress, FirstAuthorOphthoInProgress)

pc2 <- prcomp(data2, scale. = FALSE)
pc2$rotation
summary(pc2)


my_data <- my_data %>% mutate(
         PCA1 = 0.7857 * PubBeforeSubmission + 0.4925 * OphthoSubmission +
            0.30037 * FirstAuthorSubmission + 0.22333 * FirstAuthorOphthoSubmission,
         PCA2 = 0.702 * PubsInProgress + 0.550 * OphthoInProgress +
            0.3474 * FirstAuthorInProgress + 0.2895 * FirstAuthorOphthoInProgress,
         TopSchool = ifelse(TopMedSchool == 1, "Yes", "No"))
toptwentysubmission <- glm(TopDoximity ~ PCA1 + HighestIFSubmission + 
        AverageIFSubmission + Gender + TopSchool, data=my_data, family =binomial())
Table5 <- toptwentysubmission %>% tbl_regression(exponentiate = TRUE, 
        label = list(
        PCA1 ~ "ComponentAnalysisVariable",
        HighestIFSubmission ~ "Highest Impact Factor of Publications Before Submission", 
        AverageIFSubmission ~ "Average Impact Factor of Publications Before Submission", 
        TopSchool ~ "Top 40 US News Medical School: Research",
        Gender ~ "Gender (Reference = Female)"),
        digits = all_continuous() ~ 1) %>% 
        bold_labels() %>% 
        bold_p(t = 0.05) %>% modify_header(estimate = "Odds Ratio")


toptwentyInProgress <- glm(TopDoximity ~ PCA2 + 
        HighestIF + AverageIF + 
        Gender + TopSchool, data = my_data,
                family = binomial())
Table6 <- toptwentyInProgress %>% tbl_regression(exponentiate = TRUE, 
                label = list(
        PCA2 ~ "Component Analysis Variable",
        HighestIF ~ "Highest Impact Factor of all Publications",
        AverageIF ~ "Average Impact Factor of all Publications",
        TopSchool ~ "Top 40 US News Medical School: Research",
        Gender ~ "Gender (Reference = Female)"),
        digits = all_continuous() ~ 1) %>% 
                bold_labels() %>% 
                bold_p(t = 0.05) %>% modify_header(estimate = "Odds Ratio")
CombinedTableRegression <- tbl_merge (
                tbls = list(Table5, Table6), 
                tab_spanner = c("**Before Submission**", "**In Progress**")) %>% 
                gtsummary::as_flextable()

CombinedTable12 <- width(CombinedTable12, j= ~label, width=1.5)
CombinedTable12 <- width(CombinedTable12, j= ~stat_0_1+stat_1_2+stat_2_2+p.value_2, width=0.75)
CombinedTable34 <- width(CombinedTable34, j= ~label, width=1.5)
CombinedTable34 <- width(CombinedTable34, j= ~stat_1_1+stat_2_1+p.value_1+stat_1_2+stat_2_2+p.value_2, width=0.75)
CombinedTableRegression <- width(CombinedTableRegression, j= ~label, width=1.5)
CombinedTableRegression <- width(CombinedTableRegression, j= ~estimate_1+ci_1+p.value_1+estimate_2+ci_2+p.value_2, width=0.75)


doc <- read_docx() 
doc <- body_add_par(doc, value="Table 1 - Overall + Gender", style="heading 1")
doc <- body_add_flextable(doc,value= CombinedTable12) %>% body_add_break()

doc <- body_add_par(doc, value="Table 2 - TopMed + TopDoximity", style="heading 1")
doc <- body_add_flextable(doc,value= CombinedTable34) %>% body_add_break()

doc <- body_add_par(doc, value="Table 3 - Regression", style="heading 1")
doc <- body_add_flextable(doc,value= CombinedTableRegression) %>% body_add_break()

print(doc, target="C:/Chandru/CWRU/Research/CCF/OphthoAvgPubsPaper/Tables3.docx")