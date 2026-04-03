library(tidyverse)
library(magrittr)
library(readxl)
library(googlesheets4)
library(DescTools)
library(english)
library(knitr)
library(broom)
library(openxlsx)
library(gridExtra)
library(ggrepel)

library(survey)
library(survival)
library(survminer)
library(rstpm2)
library(MASS)
library(fitdistrplus)
library(MuMIn)

library(irr)
library(arules)
library(arulesViz)
library(arulesCBA)
library(tree)
library(rpart)
library(rpart.plot)
library(ggparty)
library(partykit)
library(C50)
library(caret)

options(na.action = "na.fail")
gs4_deauth()

sheets   <- read_excel("Responses.xlsx")
prereq   <- read_excel("Prereq.xlsx")
path     <-            "C:\\Users\\Owner\\Desktop\\QR Code"
analysis <-            "C:\\Users\\Owner\\Desktop\\QR Code\\Analysis.tex"
figures  <-            "C:\\Users\\Owner\\Desktop\\QR Code\\Figures\\"

sheets %<>%
  transmute(ID                         = 1:nrow(sheets),
            Timestamp                  = Timestamp,
            `Email Address`            = `Email Address`, 
            `Employment wait time`     = sheets[[40]],
            Age                        = sheets[[4]],
            Sex                        = sheets[[5]],
            `Undergraduate program`    = sheets[[6]],
            `Program preference`       = coalesce(sheets[[7]],  sheets[[12]], sheets[[17]], sheets[[22]], sheets[[27]], sheets[[32]]),
            `Graduation year`          = coalesce(sheets[[8]],  sheets[[13]], sheets[[18]], sheets[[23]], sheets[[28]], sheets[[33]]),
            `Postgraduate education`   = coalesce(sheets[[9]],  sheets[[14]], sheets[[19]], sheets[[24]], sheets[[29]], sheets[[34]]),
            `Professional exams`       = ifelse(is.na(coalesce(sheets[[10]], sheets[[15]], sheets[[20]], sheets[[25]], sheets[[30]], sheets[[35]])), 0,
                                                str_count(coalesce(sheets[[10]], sheets[[15]], sheets[[20]], sheets[[25]], sheets[[30]], sheets[[35]]), ",") + 1),
            `First employment`         = coalesce(sheets[[11]], sheets[[16]], sheets[[21]], sheets[[26]], sheets[[31]], sheets[[36]]),
            Context                    = sheets[[37]],
            `Nature of employment`     = sheets[[38]],
            `Employment qualification` = sheets[[39]],
            Status                     = ifelse(`First employment` == "I am not employed yet.", 0, 1),
            `Status string`            = ifelse(Status == 0, "Unemployed", "Employed")
            )

survey <- sheets %>%
  left_join(prereq %>% dplyr::select(c("First employment", "Sector of employment")) %>% distinct(), by = "First employment") %>%
  dplyr::select(-c(2:3, 12:13)) %>%
  relocate(last_col(), .before = 11)

workbook <- loadWorkbook("Manual.xlsx")
manual   <- read_excel("Manual.xlsx", range = cell_cols("A:H"),
                     sheet = 1)

if (nrow(sheets) > nrow(manual)){
  manual[(nrow(manual) + 1):nrow(sheets), ] <- NA
  }

manual$ID                     <- sheets$ID
manual$`First employment`     <- sheets$`First employment`
manual$Context                <- sheets$Context
manual$`Sector of employment` <- survey$`Sector of employment`

survey$`Sector of employment` <- ifelse(is.na(manual$`Sector Alternate`), survey$`Sector of employment`, manual$`Sector Alternate`)

survey %<>%
  mutate(`Undergraduate program`    = dplyr::recode(`Undergraduate program`,
                                             "BS Biology"                  = "BSB",
                                             "BS Environmental Science"    = "BSES",
                                             "BS Food Technology"          = "BSFT",
                                             "BSM Applied Statistics"      = "BSM AS",
                                             "BSM Business Applications"   = "BSM BA",
                                             "BSM Computer Science"        = "BSM CS"
                                             ),
         `Program preference`       = ifelse(`Program preference`       == "Yes", "Preferred", "Not preferred"),
         `Postgraduate education`   = ifelse(`Postgraduate education`   == "Yes", "Has plans", "No plans"),
         `Nature of employment`     = str_remove(`Nature of employment`, " —.*"),
         `Employment qualification` = ifelse(`Employment qualification` == "No specific qualification required", "None", `Employment qualification`),
         )

survey %<>%
  mutate(Sex                        = factor(Sex,                        levels = c("Male", "Female")),
         `Program preference`       = factor(`Program preference`,       levels = c("Preferred", "Not preferred")),
         `Postgraduate education`   = factor(`Postgraduate education`,   levels = c("Has plans", "No plans")),
         `Nature of employment`     = factor(`Nature of employment`,     levels = c("Regular", "Contractual", "Part-time", "Self-employed")),
         `Sector of employment`     = factor(`Sector of employment`),
         `Employment qualification` = factor(`Employment qualification`, levels = c("None", "High school diploma", "Bachelor's degree", "Postgraduate degree")),
         `Undergraduate program`    = factor(`Undergraduate program`,    levels = c("BSB", "BSES", "BSFT", "BSM AS", "BSM BA", "BSM CS"))
         )

quota <- tibble(`Undergraduate program` = survey$`Undergraduate program` %>% unique() %>% sort(),
                Quota                   = c(78, 27, 37, 40, 45, 79),
                )

surveyOld <- survey
survey %<>%
  left_join(quota, by = c("Undergraduate program")) %>%
  group_by(`Undergraduate program`) %>%
  group_modify(~ .x %>%
                 arrange(desc(`Employment wait time`), Sex, `Professional exams`) %>%
                 slice_head(n = min(nrow(.x), .x$Quota[1]))
               ) %>%
  ungroup() %>%
  dplyr::select(c(2:5, 1, 6:14)) %>%
  arrange(ID)

set.seed(13)
survey %<>%
  left_join(quota, by = c("Undergraduate program")) %>%
  group_by(`Undergraduate program`) %>%
  group_modify(~ {
    df <- .x %>%
      arrange(desc(`Employment wait time`), Sex, `Professional exams`)
    
    needed <- .x$Quota[1] - nrow(df)
    
    if (needed > 0) {
      boot <- df %>%
        slice_sample(n = needed, replace = T) %>%
        mutate(.boot = T)
      
      df %<>% mutate(.boot = F)
      df %<>% bind_rows(boot)
    } else {
      df %<>% mutate(.boot = F)
    }
    
    df
  }) %>%
  ungroup() %>%
  mutate(
    ID = if_else(.boot, 900 + cumsum(.boot), ID)
  ) %>%
  dplyr::select(-.boot) %>%
  dplyr::select(c(2:5, 1, 6:14)) %>%
  arrange(ID)

n        <- nrow(survey)

nSummary <- survey %>%
  count(`Undergraduate program`, name = "n") %>%
  left_join(quota, by = "Undergraduate program") %>%
  mutate(Count = paste0(n, "/", Quota)) %>%
  dplyr::select(`Undergraduate program`, Count)

vars <- c("Age",
          "Sex",
          "Undergraduate program",
          "Program preference",
          "Graduation year",
          "Postgraduate education",
          "Professional exams",
          "Nature of employment",
          "Sector of employment",
          "Employment qualification",
          "Status string"
          )

tibbleList <- list()
for(var in vars){
  tibbleList[[var]] <- survey %>%
    filter(!is.na(.data[[var]])) %>%
    group_by(.data[[var]]) %>%
    summarize(count = n(), .groups = "drop") %>%
    mutate(proportion = count/sum(count)) %>%
    arrange(.data[[var]])
  }

pie     <- c("Sex", 
             "Program preference", 
             "Postgraduate education", 
             "Nature of employment", 
             "Employment qualification",
             "Status string"
             )

histo   <- c("Age", 
             "Graduation year", 
             "Professional exams"
             )

colors4 <- c("steelblue", "coral", "seagreen", "firebrick")
colors6 <- c("steelblue", "coral", "seagreen", "firebrick", "mediumpurple", "gray50")









for (var in histo) {
  survey[[var]] <- as.numeric(as.character(survey[[var]]))
  plot <- ggplot(survey, aes(x = .data[[var]])) +
    geom_histogram(binwidth = 1, fill = "steelblue", color = "steelblue") +
    xlab(paste0(var)) +
    ylab("Count") +
    # scale_x_continuous(breaks = sort(unique(survey[[var]]))) +
    # scale_y_continuous(breaks = seq(0, 160, 20), limits = c(0, 160)) +
    theme_minimal() +
    theme(plot.background = element_rect(fill = "white", color = NA),
          axis.text.x     = element_text(size = 17),
          axis.text.y     = element_text(size = 17),
          axis.title      = element_text(size = 17),
          )
  # ggsave(filename = paste0(figures, "Desc ", var, ".png"),
  #        plot     = plot,
  #        width    = 6,
  #        height   = 6,
  #        dpi      = 600
  #        )
  }










ggplot(tibbleList$`Undergraduate program`, aes(x = reorder(`Undergraduate program`, count), y = count)) +
  geom_bar(stat       = "identity", fill = "steelblue") +
  geom_text(aes(label = paste0(count, " (", round(count/sum(count)*100, 0), "%)")),
            # label = paste0(count, "\n(", round(count/sum(count)*100, 0), "%)")),
            hjust     = -0.1, # -0.5
            size      = 6
            ) +
  xlab("Undergraduate program") +
  ylab("Count") +
  scale_x_discrete(labels   = function(x) str_wrap(x, width = 5)) +
  scale_y_continuous(breaks = seq(0, 100, 20), limits = c(0, 100)) +
  theme_minimal() +
  theme(plot.background = element_rect(fill = "white", color = NA),
        axis.text.x     = element_text(size = 17),
        axis.text.y     = element_text(size = 17),
        axis.title      = element_text(size = 17),
        ) +
  coord_flip()
# ggsave(filename = paste0(figures, "Desc Undergraduate program.png"),
#        width    = 6,
#        height   = 6,
#        dpi      = 600
#        )

ggplot(tibbleList$`Sector of employment`, aes(x = reorder(`Sector of employment`, -count), y = count)) +
  geom_bar(stat       = "identity", fill = "steelblue") +
  # geom_text(aes(label = paste0(count, "\n(", round(count/sum(count)*100, 0), "%)")),
  #           vjust     = -0.5,
  #           size      = 6
  #           ) +
  xlab("Sector of employment") +
  ylab("Count") +
  scale_x_discrete(labels   = function(x) str_wrap(x, width = 12)) +
  # scale_y_continuous(breaks = seq(0, 100, 20), limits = c(0, 100)) +
  theme_minimal() +
  theme(plot.background = element_rect(fill = "white", color = NA),
        axis.text.x     = element_text(size = 17),
        axis.text.y     = element_text(size = 17),
        axis.title      = element_text(size = 17),
        )
# ggsave(filename = paste0(figures, "Desc Sector of employment.png"),
#        width    = 11,
#        height   = 6,
#        dpi      = 600
#        )









for(var in pie){
  tibblePie <- tibbleList[[var]]
  original <- levels(tibblePie[[var]])
  if(is.null(original)) original <- unique(tibblePie[[var]])
  
  tibblePie[[var]] <- factor(tibblePie[[var]], 
                      levels = tibblePie[[var]][order(tibblePie$count, decreasing = TRUE)])
  
  plot <- ggplot(tibblePie, aes(x = "", y = count, fill = .data[[var]])) + 
    geom_col(width = 1, color = "white") + 
    coord_polar(theta = "y", direction = -1) + 
    xlab("") + 
    ylab("") + 
    geom_text(aes(label = paste0(count, "\n", round(count/sum(count)*100, 0), "%"), x = 1.7), 
              # label = paste0(count, "\n(", round(count/sum(count)*100, 0), "%)"), x = 1.7), 
              position = position_stack(vjust = 0.5), 
              size = 6.5, color = "black", lineheight = 0.8 
              ) + 
    scale_fill_manual(values = colors4, breaks = original) + 
    theme(plot.title = element_text(hjust = 0.5)) + 
    theme_void() + 
    theme(plot.background = element_rect(fill = "white", color = NA), 
          text = element_text(size = 21), 
          legend.position = "top", 
          legend.margin = margin(t = 0, b = -1, l = 0, r = 0, unit = "cm")
          ) + 
    labs(fill = NULL)
  # ggsave(filename = paste0(figures, "Desc ", var, ".png"), 
  #        plot   = plot, 
  #        width  = 6, 
  #        height = 6, 
  #        dpi    = 600
  #        )
}









ggplot(survey, aes(x = `Employment wait time`)) + 
  geom_histogram(binwidth = 5, fill = "steelblue", color = "steelblue") + 
  xlab("Employment wait time") + 
  ylab("Count") + 
  # scale_x_continuous(breaks = seq(0, 200, 20), limits = c(0, 200)) + 
  # scale_y_continuous(breaks = seq(0, 100, 5),  limits = c(0, 100)) + 
  theme_minimal() + 
  theme(plot.background = element_rect(fill = "white", color = NA),
        axis.text.x     = element_text(size = 17),
        axis.text.y     = element_text(size = 17),
        axis.title      = element_text(size = 17)  
        )
# ggsave(filename = paste0(figures, "Desc Employment wait time.png"),
#        width    = 6,
#        height   = 6,
#        dpi      = 600
#        )

survPlot <- survfit(Surv(`Employment wait time`, Status) ~ 1,
                    survey) %>%
  ggsurvplot(data         = survey,
             xlab         = "Employment wait time, in weeks",
             ylab         = "Probability of unemployment",
             palette      = colors4,
             legend.title = "",
             )
survPlot$plot <- survPlot$plot +
  scale_x_continuous(breaks = c(0, 6, 50, 100, 133, 140, 150)) +
  geom_hline(yintercept = 0.5, color = "black") +
  guides(color = guide_legend(nrow = 3, ncol = 2, byrow = F)) +
  theme(panel.background = element_rect(fill = "white", color = NA),
        plot.background  = element_rect(fill = "white", color = NA),
        legend.text      = element_text(size = 16),
        )
# ggsave(filename = paste0(figures, "KM No strata.png"),
#        plot     = survPlot$plot,
#        width    = 11,
#        height   = 6,
#        dpi      = 600
#        )









ageIndex <- which.max(tibbleList$Age$proportion)
ageMax   <- tibbleList$Age$Age[ageIndex]
ageSkew  <- Skew(tibbleList$Age$count)
if (is.nan(ageSkew)) ageSkew <- 0
ageKurt  <- Kurt(tibbleList$Age$count)
if (is.nan(ageKurt)) ageKurt <- 0

if (abs(ageSkew) < 0.5) {
  ageSkewIntp <- "approximately symmetric"
} else if (ageSkew >= 0.5) {
  ageSkewIntp <- "positively skewed"
} else {
  ageSkewIntp <- "negatively skewed"
}

if (abs(ageKurt) < 0.5) {
  ageKurtIntp <- "approximately mesokurtic"
} else if (ageKurt >= 0.5) {
  ageKurtIntp <- "leptokurtic"
} else {
  ageKurtIntp <- "platykurtic"
}

sexMax <- tibbleList$Sex$Sex[which.max(tibbleList$Sex$proportion)]
sexMin <- tibbleList$Sex$Sex[which.min(tibbleList$Sex$proportion)]
sexKey <- ifelse(max(tibbleList$Sex$proportion) - min(tibbleList$Sex$proportion) <= 0.10,
                 "slightly greater", "greater")

examIndex <- which.max(tibbleList$`Professional exams`$proportion)

examMax <- if (tibbleList$`Professional exams`$`Professional exams`[examIndex] == 0) {
  "did not pass (or take) any exams"
} else if (tibbleList$`Professional exams`$`Professional exams`[examIndex] == 1) {
  paste("passed", as.english(tibbleList$`Professional exams`$`Professional exams`[examIndex]), "exam")
} else {
  paste("passed", as.english(tibbleList$`Professional exams`$`Professional exams`[examIndex]), "exams")
}

natureIndex <- which.max(tibbleList$`Nature of employment`$proportion)
natureMax   <- tibbleList$`Nature of employment`$`Nature of employment`[natureIndex]

qualifIndex <- which.max(tibbleList$`Employment qualification`$proportion)
qualifMax   <- if (qualifIndex == 1) {
  "with no minimum educational qualification specified"
} else if (qualifIndex == 2) {
  "requiring a high school diploma as the minimum educational qualification"
} else if (qualifIndex == 3) {
  "requiring a Bachelor's degree as the minimum educational qualification"
} else {
  "requiring a Postgraduate degree as the minimum educational qualification"
}

timeKurtIntp <- if (abs(Kurt(survey$`Employment wait time`)) < 0.5) {
  "approximately mesokurtic"
} else if (Kurt(survey$`Employment wait time`) >= 0.5) {
  "leptokurtic"
} else {
  "platykurtic"
}

empStatusMin <- tibbleList$`Status string`$proportion[2] * 100

cat("", file = analysis)
cat(sprintf("\\newcommand{\\ageMax}{%s}", ageMax),
    sprintf("\\newcommand{\\ageSkew}{%.2f}", ageSkew),
    sprintf("\\newcommand{\\ageSkewIntp}{%s}", ageSkewIntp),
    sprintf("\\newcommand{\\ageKurt}{%.2f}", ageKurt),
    sprintf("\\newcommand{\\ageKurtIntp}{%s}", ageKurtIntp),

    sprintf("\\newcommand{\\sexMax}{%s}", tolower(sexMax)),
    sprintf("\\newcommand{\\sexMin}{%s}", tolower(sexMin)),
    sprintf("\\newcommand{\\sexKey}{%s}", sexKey),

    sprintf("\\newcommand{\\preferenceMax}{%.0f}", tibbleList$`Program preference`$proportion[1]*100),
    sprintf("\\newcommand{\\preferenceKey}{%s}", ifelse(tibbleList$`Program preference`$proportion[1] < 0.50,
                                                        "only", "about")),

    sprintf("\\newcommand{\\postgradMax}{%.0f}", tibbleList$`Postgraduate education`$proportion[1]*100),
    sprintf("\\newcommand{\\postgradKey}{%s}", ifelse(tibbleList$`Postgraduate education`$proportion[1] < 0.50,
                                                      "only", "about")),

    sprintf("\\newcommand{\\examMax}{%s}", examMax),

    sprintf("\\newcommand{\\natureMax}{%s}", tolower(natureMax)),

    sprintf("\\newcommand{\\qualifMax}{%s}", qualifMax),

    sprintf("\\newcommand{\\timeSkew}{%.2f}", Skew(survey$`Employment wait time`)),
    sprintf("\\newcommand{\\timeKurt}{%.2f}", Kurt(survey$`Employment wait time`)),
    sprintf("\\newcommand{\\timeKurtIntp}{%s}", timeKurtIntp),
    sprintf("\\newcommand{\\timeMax}{%.0f}", max(survey$`Employment wait time`)),

    sprintf("\\newcommand{\\empStatusMin}{%.0f}", empStatusMin),

    file = analysis,
    append = T,
    sep = "\n"
    )









kmVars <- c("Sex",
            "Undergraduate program",
            "Program preference",
            "Postgraduate education"
            )

for (strat in kmVars) {
  survey[[strat]] %<>% as.factor()
  kmFit <- survfit(as.formula(paste0("Surv(`Employment wait time`, Status) ~ survey$`", strat,"`")),
                   survey)
  if (!is.null(kmFit$strata)) {
    names(kmFit$strata) <- gsub(".*=\\s*", "", names(kmFit$strata))
  }
  survPlot <- ggsurvplot(kmFit,
                         data         = survey,
                         xlab         = "Employment wait time, in weeks",
                         ylab         = "Probability of unemployment",
                         palette      = colors6,
                         legend.title = "",
                         )
  survPlot$plot <- survPlot$plot + 
    guides(color = guide_legend(nrow = 3, ncol = 2, byrow = F)) +
    theme(panel.background = element_rect(fill = "white", color = NA),
          plot.background  = element_rect(fill = "white", color = NA),
          legend.text      = element_text(size = 15),
          )
  # ggsave(filename = paste0(figures, "KM ", strat, ".png"),
  #        plot     = survPlot$plot,
  #        width    = 11,
  #        height   = 6,
  #        dpi      = 600
  #        )
  }









latexWrite <- function(tbl, i, last = FALSE) {
  paste0(
    paste(
      sapply(tbl[i, ], function(x) {
        if (is.na(x)) " " else as.character(unname(x))
      }),
      collapse = " & "
    ),
    if (last) " \\\\ \\hline" else " \\\\"
  )
}

logrank <- function(x){
  as.formula(paste0("Surv(`Employment wait time`, Status) ~ survey$`", x,"`")) %>%
     survdiff(survey, rho = 0) %>%
    `[[`("chisq")
  }

logrankSig <- function(x){
  test <- as.formula(paste0("Surv(`Employment wait time`, Status) ~ survey$`", x,"`")) %>%
    survdiff(survey, rho = 0)
  1 - pchisq(test$chisq, length(test$n) - 1)
  }

gehanWilcoxon <- function(x){
  as.formula(paste0("Surv(`Employment wait time`, Status) ~ survey$`", x,"`")) %>%
    survdiff(survey, rho = 1) %>%
    `[[`("chisq")
  }

gehanWilcoxonSig <- function(x){
  test <- as.formula(paste0("Surv(`Employment wait time`, Status) ~ survey$`", x,"`")) %>%
    survdiff(survey, rho = 1)
  1 - pchisq(test$chisq, length(test$n) - 1)
  }

kmSummary   <- tibble(Variable = kmVars,
                      `Log-rank test statistic`       = sapply(Variable, logrank),
                      `Log-rank p-value`              = sapply(Variable, logrankSig),
                      `Gehan-Wilcoxon test statistic` = sapply(Variable, gehanWilcoxon),
                      `Gehan-Wilcoxon p-value`        = sapply(Variable, gehanWilcoxonSig),
                      )

kmSigN      <- sum(kmSummary$`Log-rank p-value` <= 0.05)
kmSignif    <- kmSummary %>%
  filter(`Log-rank p-value` <= 0.05)

kmSignifVar <- tolower(kmSignif$Variable)
kmSignifChi <- round(kmSignif$`Log-rank test statistic`, 2)
kmSignifSig <- kmSignif$`Log-rank p-value`
kmSignifSig <- ifelse(kmSignifSig < 0.001, "< 0.001", sprintf("%.3f", kmSignifSig))

kmLogRank <- if (kmSigN == 0) {
  "The log-rank test indicated no significant difference in employment wait time by any of the categorical variables."
} else if (kmSigN == 1) {
  paste0("The log-rank test indicated a significant difference in employment wait time only for ", 
         kmSignifVar, " ($\\chi^2 = ", kmSignifChi, "$, $p = ", kmSignifSig, "$).")
} else if (kmSigN == 2) {
  items <- paste0(kmSignifVar, " ($\\chi^2 = ", kmSignifChi, "$, $p = ", kmSignifSig, "$)")
  items_txt <- paste0(paste(items[1:(kmSigN-1)], collapse = ", "), " and ", items[kmSigN])
  paste0("The log-rank test indicated a significant difference in employment wait time by ", items_txt, ".")
} else {
  items <- paste0(kmSignifVar, " ($\\chi^2 = ", kmSignifChi, "$, $p = ", kmSignifSig, "$)")
  items_txt <- paste0(paste(items[1:(kmSigN-1)], collapse = ", "), ", and ", items[kmSigN])
  paste0("The log-rank test indicated a significant difference in employment wait time by the following: ", items_txt, ".")
}

kmSigN      <- sum(kmSummary$`Gehan-Wilcoxon p-value` <= 0.05)
kmSignif    <- kmSummary %>%
  filter(`Gehan-Wilcoxon p-value` <= 0.05)

kmSignifVar <- tolower(kmSignif$Variable)
kmSignifChi <- round(kmSignif$`Gehan-Wilcoxon test statistic`, 2)
kmSignifSig <- kmSignif$`Gehan-Wilcoxon p-value`
kmSignifSig <- ifelse(kmSignifSig < 0.001, "< 0.001", sprintf("%.3f", kmSignifSig))

kmGehanWilcoxon <- if (kmSigN == 0) {
  "the Gehan-Wilcoxon test indicated no significant difference in employment wait time by any of the categorical variables."
} else if (kmSigN == 1) {
  paste0("the Gehan-Wilcoxon test indicated a significant difference in employment wait time only for ", 
         kmSignifVar, " ($\\chi^2 = ", kmSignifChi, "$, $p = ", kmSignifSig, "$).")
} else if (kmSigN == 2) {
  items <- paste0(kmSignifVar, " ($\\chi^2 = ", kmSignifChi, "$, $p = ", kmSignifSig, "$)")
  itemsText <- paste0(paste(items[1:(kmSigN-1)], collapse = ", "), " and ", items[kmSigN])
  paste0("the Gehan-Wilcoxon test indicated a significant difference in employment wait time by ", itemsText, ".")
} else {
  items <- paste0(kmSignifVar, " ($\\chi^2 = ", kmSignifChi, "$, $p = ", kmSignifSig, "$)")
  itemsText <- paste0(paste(items[1:(kmSigN-1)], collapse = ", "), ", and ", items[kmSigN])
  paste0("the Gehan-Wilcoxon test indicated a significant difference in employment wait time by the following: ", itemsText, ".")
}

cat(sprintf("\\newcommand{\\kmLogRank}{%s}", kmLogRank),
    sprintf("\\newcommand{\\kmGehanWilcoxon}{%s}", kmGehanWilcoxon),
    file = analysis,
    append = T,
    sep = "\n"
    )

kmSummaryAlt <- kmSummary %>%
  mutate(Variable                        = kmVars,
         `Log-rank test statistic`       = `Log-rank test statistic`       %>% round(2) %>% format(2),
         `Log-rank p-value`              = ifelse(`Log-rank p-value` < 0.001, "<0.001", format(round(`Log-rank p-value`, 3), nsmall = 3)),
         `Gehan-Wilcoxon test statistic` = `Gehan-Wilcoxon test statistic` %>% round(2) %>% format(2),
         `Gehan-Wilcoxon p-value`        = ifelse(`Gehan-Wilcoxon p-value` < 0.001, "<0.001", format(round(`Gehan-Wilcoxon p-value`, 3), nsmall = 3)),
         )

sapply(1:nrow(kmSummaryAlt), function(i) latexWrite(kmSummaryAlt, i, last = i == nrow(kmSummaryAlt)))# %>%
#  writeLines("Chapter 4 KM Summary.tex")









timeVars <- c("Age", "Sex", "Undergraduate program", "Program preference", 
              "Graduation year", "Postgraduate education", "Professional exams"
              )

surveyDelete <- survey %>%
  dplyr::select(-c(1, 10:12, 14)) %>%
  mutate(`Graduation year` = `Graduation year` - 2022,
         Age               = Age - 22)

coxFit <- coxph(Surv(surveyDelete$`Employment wait time`, Status) ~ .,
                surveyDelete)

coxAssump <- cox.zph(coxFit)$table %>% as_tibble() %>%
  mutate(Variable   = c(timeVars, "Global"),
         Assumption = ifelse(p > 0.05, "Yes", "No")
         )
coxAssumpAlt <- coxAssump %>%
  mutate(chisq = chisq %>% round(2) %>% format(2),
         p     = ifelse(p < 0.001, "<0.001", format(round(p, 3), nsmall = 3))
         ) %>%
  dplyr::select(Variable, everything())


sapply(1:nrow(coxAssumpAlt), function(i) latexWrite(coxAssumpAlt, i, last = i == nrow(coxAssumpAlt)))# %>%
#  writeLines("Chapter 4 Cox Assumption.tex")


coxAssump %<>%
  filter(Assumption == "Yes")
coxString <- coxAssump$Variable %>%
  setdiff("Global") %>%
  paste0("`", ., "`") %>%
  paste(collapse = " + ")
coxSummary <- coxph(as.formula(paste("Surv(`Employment wait time`, Status) ~ ", coxString)),
                    surveyDelete
                    ) %>% summary()
coxSummary <- tibble(Variable = coxSummary$coefficients %>% rownames(),
                     exp      = coxSummary$coefficients[, "exp(coef)"],
                     se       = coxSummary$coefficients[, "se(coef)"] ,
                     z        = coxSummary$coefficients[, "z"],
                     p        = coxSummary$coefficients[, "Pr(>|z|)"],
                     )
coxSummaryAlt <- coxSummary %>%
  mutate(exp      = exp      %>% round(2) %>% format(2),
         se       = se       %>% round(2) %>% format(2),
         z        = z        %>% round(2) %>% format(2),
         p        = ifelse(p < 0.001, "<0.001", format(round(p, 3), nsmall = 3))
         )

sapply(1:nrow(coxSummaryAlt), function(i) latexWrite(coxSummaryAlt, i, last = i == nrow(coxSummaryAlt)))# %>%
#  writeLines("Chapter 4 Cox Summary.tex")

coxBest <- stepAIC(coxFit, direction = "both") %>% tidy() %>% mutate(expo = exp(estimate)) %>%
  dplyr::select(1, 2, 6, 3, 5)
coxBestAlt <- coxBest %>%
  mutate(estimate  = estimate  %>% round(2) %>% format(2),
         expo      = expo      %>% round(2) %>% format(2),
         std.error = std.error %>% round(2) %>% format(2),
         p.value = ifelse(p.value < 0.001, "<0.001", format(round(p.value, 3), nsmall = 3))
         )

sapply(1:nrow(coxBestAlt), function(i) latexWrite(coxBestAlt, i, last = i == nrow(coxBestAlt)))# %>%
#  writeLines("Chapter 4 Cox Best.tex")


coxM <- stepAIC(coxFit, direction = "both")
mart <- residuals(coxM, type = "martingale")
contVars <- c("Age", "Graduation year", "Professional exams")

for (var in contVars) {
  lowessFit <- lowess(surveyDelete[[var]], mart)
  lowessTibble <- data.frame(x = lowessFit$x, y = lowessFit$y)
  
  plot <- ggplot(surveyDelete, aes(x = .data[[var]], y = mart)) +
    geom_point(position = position_jitter(width = 0.2), alpha = 0.6, color = "steelblue") +
    geom_hline(yintercept = 0, color = "coral") +
    geom_smooth(method = "loess", se = FALSE, color = "black", span = 1) +
    # geom_line(data = lowessTibble, aes(x = x, y = y), color = "black", size = 1) +
    labs(x = var, y = "Martingale residuals") +
    theme(panel.background = element_rect(fill = "white", color = NA),
          plot.background  = element_rect(fill = "white", color = NA),
          )
  # ggsave(filename = paste0(figures, "Cox Martingale ", var, ".png"),
  #        plot   = plot,
  #        width  = 11,
  #        height = 6,
  #        dpi    = 600
  #        )
}









surveyDelete$`Employment wait time` <- surveyDelete$`Employment wait time` + 0.1

aftExp    <- survreg(Surv(`Employment wait time`, Status) ~ ., surveyDelete, dist = "exponential")
aftGau    <- survreg(Surv(`Employment wait time`, Status) ~ ., surveyDelete, dist = "gaussian")
aftLgs    <- survreg(Surv(`Employment wait time`, Status) ~ ., surveyDelete, dist = "logistic")
aftLogLgs <- survreg(Surv(`Employment wait time`, Status) ~ ., surveyDelete, dist = "loglogistic")
aftLogNor <- survreg(Surv(`Employment wait time`, Status) ~ ., surveyDelete, dist = "lognormal")
aftWei    <- survreg(Surv(`Employment wait time`, Status) ~ ., surveyDelete, dist = "weibull")

aft <- list(Exponential = aftExp, 
            Gaussian    = aftGau,
            Logistic    = aftLgs, 
            Loglogistic = aftLogLgs, 
            Lognormal   = aftLogNor, 
            Weibull     = aftWei
            )

aftSelect <- tibble(Distribution = names(aft),
                    chisq = sapply(aft, function(x) summary(x)$chi),
                    p     = sapply(aft, function(x) pchisq(summary(x)$chi, summary(x)$df - 1, lower.tail = F)),
                    AIC   = sapply(aft, AIC),
                    BIC   = sapply(aft, BIC),
                    ) 

aftSelectAlt <- aftSelect %>%
  mutate(chisq = chisq %>% round(2) %>% format(2),
         p     = ifelse(p < 0.001, "<0.001", format(round(p, 3), nsmall = 3)),
         AIC   = AIC   %>% round(2) %>% format(2),
         BIC   = BIC   %>% round(2) %>% format(2),
         )
sapply(1:nrow(aftSelectAlt), function(i) latexWrite(aftSelectAlt, i, last = i == nrow(aftSelectAlt)))# %>%
#  writeLines("Chapter 4 AFT Select.tex")

aftLogLgs0 <- survreg(Surv(`Employment wait time`, Status) ~ 1,
                      data = surveyDelete,
                      dist = "loglogistic"
                      )

logLgsP <- function(x, shape, scale) {
  1 / (1 + (x / scale)^(-shape))
}

logLgsShape <- 1 / aftLogLgs0$scale
logLgsScale <- exp(aftLogLgs0$coefficients)

timeSorted  <- sort(surveyDelete$`Employment wait time`)
timeEmp     <- ecdf(surveyDelete$`Employment wait time`)

aftFitDist  <- list(Exponential = fitdist(timeSorted, "exp"),
                    Gaussian    = fitdist(timeSorted, "norm"),
                    Lognormal   = fitdist(timeSorted, "lnorm"),
                    Logistic    = fitdist(timeSorted, "logis"),
                    Weibull     = fitdist(timeSorted, "weibull")
                    )

tibble(Time        = timeSorted,
       Empirical   = timeEmp(timeSorted),
       Exponential = pexp(timeSorted, rate = aftFitDist$Exponential$estimate["rate"]),
       Gaussian    = pnorm(timeSorted,
                           mean = aftFitDist$Gaussian$estimate["mean"],
                           sd   = aftFitDist$Gaussian$estimate["sd"]
                           ),
       Logistic    = plogis(timeSorted,
                            location = aftFitDist$Logistic$estimate["location"],
                            scale    = aftFitDist$Logistic$estimate["scale"]
                            ),
       Loglogistic = logLgsP(timeSorted,
                             shape = logLgsShape,
                             scale = logLgsScale
                             ),
       Lognormal   = plnorm(timeSorted,
                            meanlog = aftFitDist$Lognormal$estimate["meanlog"],
                            sdlog   = aftFitDist$Lognormal$estimate["sdlog"]
                            ),
       Weibull     = pweibull(timeSorted,
                              shape = aftFitDist$Weibull$estimate["shape"],
                              scale = aftFitDist$Weibull$estimate["scale"]
                              )
       ) %>%
  pivot_longer(cols = c(Exponential, Gaussian, Logistic, Loglogistic, Lognormal, Weibull),
               names_to = "Distribution",
               values_to = "Theoretical"
               ) %>%
  ggplot(aes(Theoretical, Empirical, color = Distribution)) +
  geom_abline(slope = 1, intercept = 0) +
  geom_point(size = 3) +
  labs(x = "Theoretical probabilities",
       y = "Empirical probabilities"
       ) +
  theme_minimal() +
  theme(plot.background   = element_rect(fill = "white", color = NA),
        axis.text         = element_text(size = 16),
        axis.title        = element_text(size = 16),
        legend.text       = element_text(size = 16),
        legend.title      = element_text(size = 16),
        legend.key.height = unit(1.2, "cm")
        ) +
  guides(color = guide_legend(override.aes = list(size = 4))) +
  scale_color_manual(values = colors6)
# ggsave(filename = paste0(figures, "AFT PP Plot.png"),
#        width    = 11,
#        height   = 6,
#        dpi      = 600
#        )










aftIndex   <- which.min(aftSelect$BIC)

aftBest    <- stepAIC(aft[[aftIndex]], direction = "both") %>% tidy()
aftBestAlt <- aftBest %>%
  mutate(estimate = estimate   %>% round(2) %>% format(2),
         std.error = std.error %>% round(2) %>% format(2),
         statistic = statistic %>% round(2) %>% format(2),
         p.value = ifelse(p.value < 0.001, "<0.001", format(round(p.value, 3), nsmall = 3)),
         )

sapply(1:nrow(aftBestAlt), function(i) latexWrite(aftBestAlt, i, last = i == nrow(aftBestAlt)))# %>%
#  writeLines("Chapter 4 AFT Best.tex")

aftBestDeviance <- data.frame(Fitted = fitted(aft[[aftIndex]]),
                              Resid  = residuals(aft[[aftIndex]], type = "deviance"),
                              Label  = 1:n
                              )

ggplot(aftBestDeviance, aes(x = Fitted, y = Resid)) +
  geom_hline(yintercept = 0) +
  geom_point(color = "steelblue", size = 3) +
  # geom_text_repel(aes(label = Label), size = 3, max.overlaps = 20) +
  labs(x = "Fitted values",
       y = "Deviance residuals",
       ) +
  scale_y_continuous(breaks = -4:4,
                     limits = c(-4,4)
                     ) +
  theme_minimal() +
  theme(plot.background = element_rect(fill = "white", color = NA),
        axis.text.x     = element_text(size = 16),
        axis.text.y     = element_text(size = 16),
        axis.title      = element_text(size = 16),
        )
# ggsave(filename = paste0(figures, "AFT Deviance Residuals Best.png"),
#        width    = 11,
#        height   = 6,
#        dpi      = 600
#        )

printPlot <- list()
for (m in 1:6) { enterString <- ifelse(m > 3, "\n\n", "")
  printPlot[[m]] <- ggplot(data.frame(Fitted = fitted(aft[[m]]),
                                      Resid  = residuals(aft[[m]], type = "deviance")
                                      ), aes(x = Fitted, y = Resid)) +
  geom_hline(yintercept = 0) +
  geom_point(color = "steelblue", size = 3) +
  labs(x = "Fitted values",
       y = "Deviance residuals",
       title = paste0(enterString, toupper(substring(aft[[m]]$dist, 1, 1)), substring(aft[[m]]$dist, 2))
  ) +
  # scale_y_continuous(breaks = -5:5,
  #                    limits = c(-5,5)
  #                    ) +
  theme_minimal() +
  theme(plot.background = element_rect(fill = "white", color = NA),
        # axis.text.x     = element_text(size = 16),
        # axis.text.y     = element_text(size = 16),
        # axis.title      = element_text(size = 16),
        plot.title      = element_text(hjust = 0.5, size = 16,)
  )
}
grid.arrange(grobs = printPlot, nrow = 2, ncol = 3)
# png(filename = paste0(figures, "AFT Deviance Residuals.png"), 
#     width    = 2500, 
#     height   = 1800, 
#     res      = 150
#     )
grid.arrange(grobs = printPlot, nrow = 2, ncol = 3)
dev.off()

cat(sprintf("\\newcommand{\\aftBest}{%s}", if(aftIndex != 2 & aftIndex != 6){
  aftSelect$Distribution[[aftIndex]] %>% tolower()
  } else {
    aftSelect$Distribution[[aftIndex]]
  }),
  sprintf("\\newcommand{\\aftBestKey}{%s}", ifelse((aftIndex == 1), "an", "a")),
  file = analysis,
  append = T,
  sep = "\n"
  )


aftM <- stepAIC(aft[[aftIndex]], direction = "both")
aftRes <- residuals(aftM, type = "response")

for (var in contVars) {
  plot <- ggplot(surveyDelete, aes(x = .data[[var]], y = aftRes)) +
    geom_point(position = position_jitter(width = 0.2), alpha = 0.6, color = "steelblue") +
    geom_hline(yintercept = 0, color = "coral") +
    geom_smooth(method = "loess", se = FALSE, color = "black", span = 1) +
    labs(x = var, y = "Residuals") +
    theme(panel.background = element_rect(fill = "white", color = NA),
          plot.background  = element_rect(fill = "white", color = NA),
          )
  
  # ggsave(filename = paste0(figures, "AFT Residuals ", var, ".png"),
  #        plot   = plot,
  #        width  = 11,
  #        height = 6,
  #        dpi    = 600)
}









Knots <- 1:6
rpsFullFormula <- as.formula(paste("Surv(`Employment wait time`, Status) ~", 
                                   paste(paste0("`", timeVars, "`"), collapse = " + "))
                             )

rpsSelect <- tibble(Knots = Knots,
                    Model = lapply(Knots, function(x) {stpm2(rpsFullFormula, surveyDelete, df = x)}),
                    AIC   = sapply(Model, AIC),
                    BIC   = sapply(Model, BIC),
                    ) %>%
  dplyr::select(-Model)

rpsSelectAlt <- rpsSelect %>%
  mutate(AIC = AIC %>% round(2) %>% format(2),
         BIC = BIC %>% round(2) %>% format(2),
         )

sapply(1:nrow(rpsSelectAlt), function(i) latexWrite(rpsSelectAlt, i, last = i == nrow(rpsSelectAlt)))# %>%
#  writeLines("Chapter 4 RPS Select.tex")

rpsIndex <- which.min(rpsSelect$BIC)
cat(sprintf("\\newcommand{\\rpsSelect}{%s}", rpsIndex %>% as.english()),
    file = analysis,
    append = T,
    sep = "\n"
    )

rpsCurrent    <- stpm2(rpsFullFormula, surveyDelete, df = rpsIndex)
rpsCurrentAIC <- AIC(rpsCurrent)
improved      <- T

while(improved && length(timeVars) > 1) {
  rpsAIC <- numeric(length(timeVars))
  
  for(i in seq_along(timeVars)) {
    rpsTempVars <- timeVars[-i]
    rpsTempFormula <- as.formula(
      paste("Surv(`Employment wait time`, Status) ~", 
            paste(paste0("`", rpsTempVars, "`"), collapse = " + "))
    )
    rpsTemp <- stpm2(rpsTempFormula, surveyDelete, df = rpsIndex)
    rpsAIC[i] <- AIC(rpsTemp)
    }
  
  rpsBestAIC <- min(rpsAIC)
  if(rpsBestAIC < rpsCurrentAIC) {
    rpsRemoveIndex <- which.min(rpsAIC)
    cat("Dropping:", timeVars[rpsRemoveIndex], "AIC improved from", rpsCurrentAIC, "to", rpsBestAIC, "\n")
    timeVars <- timeVars[-rpsRemoveIndex]
    rpsCurrentAIC <- rpsBestAIC
    rpsCurrent <- stpm2(as.formula(
      paste("Surv(`Employment wait time`, Status) ~", 
            paste(paste0("`", timeVars, "`"), collapse = " + "))
      ), surveyDelete, df = rpsIndex)
    } else {
      improved <- F
      }
  }

cat("Final timeVars:", paste(timeVars, collapse = ", "), "\n")

rpsBest <- rpsCurrent %>% tidy()
rpsBestAlt <- rpsBest %>%
  mutate(estimate  = estimate  %>% round(2) %>% format(2),
         std.error = std.error %>% round(2) %>% format(2),
         statistic = statistic %>% round(2) %>% format(2),
         p.value   = ifelse(p.value < 0.001, "<0.001", format(round(p.value, 3), nsmall = 3)),
         )

sapply(1:nrow(rpsBestAlt), function(i) latexWrite(rpsBestAlt, i, last = i == nrow(rpsBestAlt)))# %>%
#  writeLines("Chapter 4 RPS Best.tex")









rpsM <- rpsCurrent

best <- tibble(Model = c("Cox Regression", "Weibull AFT", "Royston--Parmar Splines"),
               AIC   = c(AIC(coxM), AIC(aftM), AIC(rpsM)),
               BIC   = c(BIC(coxM), BIC(aftM), BIC(rpsM)),
               CIC   = c(summary(coxM)$concordance[1], 
                         concordance(Surv(`Employment wait time`, Status) ~ predict(aftM, type = "quantile", p = 0.5), surveyDelete)$concordance, 
                         concordance(Surv(`Employment wait time`, Status) ~ predict(rpsM, surveyDelete, "link"), surveyDelete)$concordance
                         ),
               DIC   = c(logLik(coxM), logLik(aftM), -rpsM@min),
               )

bestAlt <- best %>%
  mutate(AIC = AIC %>% round(2) %>% format(2),
         BIC = BIC %>% round(2) %>% format(2),
         CIC = CIC %>% round(2) %>% format(2),
         DIC = DIC %>% round(3) %>% format(3),
         )
sapply(1:nrow(bestAlt), function(i) latexWrite(bestAlt, i, last = i == nrow(bestAlt)))# %>%
#  writeLines("Chapter 4 Best.tex")









newdata <- data.frame(
  `Employment wait time` = 16,   
  `Age` = 0,                    
  `Sex` = "Male",
  `Undergraduate program` = "BSM CS",
  `Graduation year` = 4,
  `Professional exams` = 0,
  check.names = FALSE
  )

rstpm2::predict(rpsM, newdata = newdata, type = "rmst", se.fit = TRUE)









codebook <- sheets %>%
  mutate(`Undergraduate program` = dplyr::recode(`Undergraduate program`,
                                   "BS Biology"                  = "BSB",
                                   "BS Environmental Science"    = "BSES",
                                   "BS Food Technology"          = "BSFT",
                                   "BSM Applied Statistics"      = "BSM AS",
                                   "BSM Business Applications"   = "BSM BA",
                                   "BSM Computer Science"        = "BSM CS"
                                   )) %>%
  left_join(prereq %>% dplyr::select(c("First employment", "Program", "Sector of employment", "Relevance")), by = c("First employment", "Undergraduate program" = "Program")) %>%
  dplyr::select(-c(2:4, 12:13, 16:17)) %>%
  relocate(`Sector of employment`, .before = 10)

codebook %<>%
  mutate(`Program preference`       = ifelse(`Program preference`       == "Yes", "Preferred", "Not preferred"),
         `Postgraduate education`   = ifelse(`Postgraduate education`   == "Yes", "Has plans", "No plans"),
         `Nature of employment`     = str_remove(`Nature of employment`, " —.*"),
         `Employment qualification` = ifelse(`Employment qualification` == "No specific qualification required", "None", `Employment qualification`),
         )

manual$Program   <- codebook$`Undergraduate program`
manual$Relevance <- codebook$Relevance
  
addStyle(workbook, 
         sheet = 1, 
         style = createStyle(wrapText = T), 
         cols = 2:5, 
         rows = 1:(nrow(manual) + 1),
         gridExpand = T, stack = T
         )

addStyle(workbook, 
         sheet = 1, 
         style = createStyle(halign = "left", valign = "top"),
         cols = 1:50,
         row = 2:400,
         gridExpand = T, stack = T
         )

addStyle(workbook, 
         sheet = 1, 
         style = createStyle(halign = "center", valign = "top"),
         cols = 1:50,
         row = 1,
         gridExpand = T, stack = T
         )

addStyle(workbook, 
         sheet = 1, 
         style = createStyle(halign = "center", valign = "top"),
         cols = c(1, 4:50),
         row = 2:400,
         gridExpand = T, stack = T
         )

setColWidths(workbook, sheet = 1, cols = 1,   widths = 5)
setColWidths(workbook, sheet = 1, cols = 2,   widths = 20)
setColWidths(workbook, sheet = 1, cols = 3,   widths = 100)
setColWidths(workbook, sheet = 1, cols = 4:8, widths = 13)

writeData(workbook, sheet = "Sheet 1", x = manual)
saveWorkbook(workbook, "Manual.xlsx", overwrite = T)

codebook$`Sector of employment` <- surveyOld$`Sector of employment`
codebook$Relevance              <- ifelse(is.na(manual$`Relevance Alternate`), codebook$Relevance, manual$`Relevance Alternate`)

codebook$Relevance %<>% factor(levels = c("Highly", "Partly", "Not"), ordered = T)
codebook %<>% right_join(tibble(ID = survey$ID), by = "ID") %>%
  filter(!is.na(`Nature of employment`)) %>% dplyr::select(-1)

names(codebook) <- c("Age", "Sex", "Program", "Prefer", "Year", "Postgrad",
                     "Exams", "Nature", "Sector", "Qualif", "Relevance"
                     )









codebookID <- manual %>%
  dplyr::select(c(1, 7, 8)) %>%
  mutate(Relevance = ifelse(is.na(Relevance), `Relevance Alternate`, Relevance)) %>%
  dplyr::select(-3)

clemente <- read_excel("Clemente.xlsx", skip = 2) %>%
  dplyr::select(-(2:4))

mandap <- read_excel("Mandap.xlsx", skip = 2) %>%
  dplyr::select(-(2:4))

valeroso <- read_excel("Valeroso.xlsx", skip = 2) %>%
  dplyr::select(-(2:4))

interRater <- codebookID %>%
  right_join(clemente, by = "ID") %>%
  right_join(mandap,   by = "ID") %>%
  right_join(valeroso, by = "ID") %>%
  dplyr::select(-1)

interRater[] <- lapply(interRater, function(x) 
  factor(x, levels = c("Not", "Partly", "Highly"), 
         ordered = TRUE)
  )

rater1 <- kappa2(interRater[, c(1, 2)], weight = c(1, 3, 4))
rater2 <- kappa2(interRater[, c(1, 3)], weight = c(1, 3, 4))
rater3 <- kappa2(interRater[, c(1, 4)], weight = c(1, 3, 4))

cohens <- tibble(Rater = c("Rater 1", "Rater 2", "Rater 3"),
                 k     = c(rater1$value,     rater2$value,     rater3$value),
                 z     = c(rater1$statistic, rater2$statistic, rater3$statistic),
                 p     = c(rater1$p.value,   rater2$p.value,   rater3$p.value),
                 i     = "Moderate, significantly greater than chance")

cohensAlt <- cohens %>%
  mutate(k = k %>% round(2) %>% format(2),
         z = z %>% round(2) %>% format(2),
         p = ifelse(p < 0.001, "<0.001", format(round(p, 3), nsmall = 3)),
         )

sapply(1:nrow(cohensAlt), function(i) latexWrite(cohensAlt, i, last = i == nrow(cohensAlt)))# %>%
#  writeLines("Chapter 4 Cohen's kappa.tex")








codebookFactor <- codebook %>%
  dplyr::mutate(across(everything(), as.factor))

armFit <- CBA(Relevance ~ ., codebookFactor,
              supp = 0.10,          
              conf = 0.70            
              )

armSummary <- armFit$rules %>% head(10) %>% 
  as(., "data.frame") %>% as_tibble() %>%
  dplyr::select(c(1:3, 5)) %>%
  separate(`rules`, c("LHS", "RHS"), sep = " => ") %>%
  mutate(LHS = str_replace_all(LHS, "\\b\\w+=([^,}]+)", "\\1"),
         RHS = str_replace_all(RHS, "\\b\\w+=([^,}]+)", "\\1"),
         LHS = str_replace_all(LHS, ",", ", "),
         RHS = str_replace_all(RHS, ",", ", "),
         )

armSummaryAlt <- armSummary %>% 
  mutate(support    = support    %>% round(2) %>% format(2),
         confidence = confidence %>% round(2) %>% format(2),
         lift       = lift       %>% round(2) %>% format(2),
         )

sapply(1:nrow(armSummaryAlt), function(i) latexWrite(armSummaryAlt, i, last = i == nrow(armSummaryAlt)))# %>%
#  writeLines("Chapter 4 ARM Summary.tex")

armCM <- predict(armFit, codebookFactor) %>%
  confusionMatrix(codebookFactor$Relevance, .)

armCM$byClass %>% as_tibble() %>%
  mutate(`F1 Score` = 2/(1/`Precision` + 1/`Recall`)) %>%
  dplyr::select(c(5, 1, 2, 11, 12)) %>%
  mutate(Level  = c("Highly", "Partly", "Not")) %>%
  pivot_longer(1:5,
               names_to  = "Metric",
               values_to = "ARM") %>%
  mutate(Level  = factor(Level,  levels = c("Highly", "Partly", "Not")),
         Metric = factor(Metric, levels = c("Overall", "Precision", "Sensitivity", "Specificity", "Balanced Accuracy", "F1 Score"))) %>%
  arrange(Metric, Level) %>%
  dplyr::select(c(2, 1, 3)) %>%
  add_row(Metric = "Accuracy",      Level = "Overall", ARM = armCM$overall[[1]], .before = 1) %>%
  add_row(Metric = "Cohen's Kappa", Level = "Overall", ARM = armCM$overall[[2]], .before = 2) -> armTable









cartFit <- rpart(Relevance ~ ., data = codebook, method = "class",
                 # control = rpart.control(minsplit = round(0.10*n, 0), minbucket = 1, cp = 0.01)
                 )

# png(filename = paste0(figures, "Relevance CART.png"), 
#     width    = 2500, 
#     height   = 1800, 
#     res      = 150
#     )

rpart.plot(cartFit,
           type          = 1,
           extra         = 100,
           box.palette   = list("seagreen", "steelblue", "coral"),
           uniform       = T,
           fallen.leaves = T,
           cex           = 2,
           branch.lwd    = 2,
           gap           = 1,
           )
dev.off()

cartCM <- predict(cartFit, codebook) %>%
  apply(1, function(x) colnames(.)[which.max(x)]) %>%
  factor(levels = c("Highly", "Partly", "Not")) %>%
  confusionMatrix(codebookFactor$Relevance, .)

cartCM$byClass %>% as_tibble() %>%
  mutate(`F1 Score` = 2/(1/`Precision` + 1/`Recall`)) %>%
  dplyr::select(c(5, 1, 2, 11, 12)) %>%
  mutate(Level  = c("Highly", "Partly", "Not")) %>%
  pivot_longer(1:5,
               names_to  = "Metric",
               values_to = "CART") %>%
  mutate(Level  = factor(Level,  levels = c("Highly", "Partly", "Not")),
         Metric = factor(Metric, levels = c("Overall", "Precision", "Sensitivity", "Specificity", "Balanced Accuracy", "F1 Score"))) %>%
  arrange(Metric, Level) %>%
  dplyr::select(c(2, 1, 3)) %>%
  add_row(Metric = "Accuracy",      Level = "Overall", CART = cartCM$overall[[1]], .before = 1) %>%
  add_row(Metric = "Cohen's Kappa", Level = "Overall", CART = cartCM$overall[[2]], .before = 2) -> cartTable









codebookFactor %<>%
  mutate(Program = dplyr::recode(Program,
                                 "BSB"    = "Bio",
                                 "BSES"   = "ES",
                                 "BSFT"   = "FT",
                                 "BSM AS" = "AS",
                                 "BSM BA" = "BA",
                                 "BSM CS" = "CS"
                                 ))
c5.0Fit <- C5.0(Relevance ~ ., codebookFactor)
summary(c5.0Fit)

c5.0CM <- predict(c5.0Fit, codebookFactor) %>%
  confusionMatrix(codebook$Relevance, .)

c5.0CM$byClass %>% as_tibble() %>%
  mutate(`F1 Score` = 2/(1/`Precision` + 1/`Recall`)) %>%
  dplyr::select(c(5, 1, 2, 11, 12)) %>%
  mutate(Level  = c("Highly", "Partly", "Not")) %>%
  pivot_longer(1:5,
               names_to  = "Metric",
               values_to = "C5.0") %>%
  mutate(Level  = factor(Level,  levels = c("Highly", "Partly", "Not")),
         Metric = factor(Metric, levels = c("Overall", "Precision", "Sensitivity", "Specificity", "Balanced Accuracy", "F1 Score"))) %>%
  arrange(Metric, Level) %>%
  dplyr::select(c(2, 1, 3)) %>%
  add_row(Metric = "Accuracy",      Level = "Overall", C5.0 = c5.0CM$overall[[1]], .before = 1) %>%
  add_row(Metric = "Cohen's Kappa", Level = "Overall", C5.0 = c5.0CM$overall[[2]], .before = 2) -> c5.0Table

# png(filename = paste0(figures, "Relevance C5.0.png"), 
#     width    = 3000, 
#     height   = 1800, 
#     res      = 150
#     )
plot(as.party(c5.0Fit),
     gp = gpar(fontsize = 22, fontface = "bold"),
     )
dev.off()









allCM <- armTable %>%
  left_join(cartTable, by = c("Metric", "Level")) %>%
  left_join(c5.0Table, by = c("Metric", "Level")) %>%
  mutate(Metric  = replace(Metric, c(4, 5, 7, 8, 10, 11, 13, 14, 16, 17), NA))

allCMAlt <- allCM %>%
  mutate(ARM_num  = ARM,
         CART_num = CART,
         C5_num   = C5.0
         ) %>% rowwise() %>%
  mutate(max_val = max(c(ARM_num, CART_num, C5_num), na.rm = TRUE),
         ARM     = ifelse(!is.na(ARM_num) & ARM_num == max_val, paste0("\\textcolor{blue}{", format(round(ARM_num,3), nsmall=3), "}"), 
                          ifelse(is.na(ARM_num), "-", format(round(ARM_num,3), nsmall=3))),
         CART    = ifelse(!is.na(CART_num) & CART_num == max_val, paste0("\\textcolor{blue}{", format(round(CART_num,3), nsmall=3), "}"), 
                          ifelse(is.na(CART_num), "-", format(round(CART_num,3), nsmall=3))),
         C5.0    = ifelse(!is.na(C5_num) & C5_num == max_val, paste0("\\textcolor{blue}{", format(round(C5_num,3), nsmall=3), "}"), 
                          ifelse(is.na(C5_num), "-", format(round(C5_num,3), nsmall=3)))
         ) %>%
  ungroup() %>%
  dplyr::select(-ARM_num, -CART_num, -C5_num, -max_val)

sapply(1:nrow(allCMAlt), function(i) latexWrite(allCMAlt, i, last = i == nrow(allCMAlt)))# %>%
#  writeLines("Chapter 4 Worst.tex")









library(fpp3)
sheets %>%
  mutate(Date = as_date(Timestamp)) %>%
  group_by(Date) %>%
  summarise(Count = n()) %>%
  as_tsibble(index = Date) %>%
  fill_gaps(Count = 0) %>%
  autoplot(Count)
# ggsave(filename = paste0(figures, "Survey Answers.png"),
#        width    = 11,
#        height   = 6,
#        dpi      = 600
#        )

library(writexl)
# write_xlsx(survey, "Export.xlsx")

table(codebookFactor$Program, codebookFactor$Relevance)
