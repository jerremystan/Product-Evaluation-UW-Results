library(MASS)
library(STAT)
library(mgcv)
library(actuar)
library(insuranceData)
library(ggplot2)
library(dplyr)
library(readxl)
library(janitor)

data_frequency <- read_excel("GLM Claim Data - Frequency.xlsx", sheet = "To R")
frequency <- clean_names(data_frequency)

data_severity <- read_excel("GLM Claim Data - Severity.xlsx", sheet = "To R")
severity <- clean_names(data_severity)

as_factor_cols <- c("occupation_code","vehicle_type_code","excess_type","gender","is_aws")
freq_data <- frequency %>%
  mutate(across(all_of(as_factor_cols), ~as.factor(.)))

sev_data <- severity %>%
  mutate(across(all_of(as_factor_cols), ~as.factor(.)))

align_factor_levels <- function(newdata, traindata, cols) {
  for (nm in cols) {
    # kalau kolom belum factor di newdata, jadikan factor
    if (!is.factor(newdata[[nm]])) newdata[[nm]] <- as.factor(newdata[[nm]])
    # samakan level dengan training
    newdata[[nm]] <- factor(newdata[[nm]], levels = levels(traindata[[nm]]))
  }
  newdata
}

model_stats <- function(m) {
  pear <- residuals(m, type="pearson")
  devr <- residuals(m, type="deviance")
  
  nd <- if (!is.null(m$null.deviance)) m$null.deviance else summary(m)$null.deviance
  
  data.frame(
    AIC = AIC(m),
    BIC = BIC(m),
    LogLik = as.numeric(logLik(m)),
    Dispersion = sum(pear^2) / df.residual(m),
    PseudoR2 = 1 - (deviance(m) / nd),
    MeanPearson = mean(pear, na.rm=TRUE),
    SdPearson   = sd(pear, na.rm=TRUE),
    MaxDeviance = max(abs(devr), na.rm=TRUE)
  )
}

#FREQUENCY

#Model 1 POISSON
model_freq <- glm(
  sum_of_claim_count ~ occupation_code + vehicle_type_code + manufacturing_year + excess_type + gender + is_aws,
  family = poisson(link = "log"),
  offset = log(exposure),
  data = freq_data
)

summary(model_freq)

# m_full <- model_freq
# m_reduced <- update(m_full, . ~ . - gender)
# anova(m_reduced, m_full, test = "Chisq")

# #Model 2 POISSON
# model_freq2 <- glm(
#   sum_of_claim_count ~ occupation_code + vehicle_type_code + manufacturing_year + gender + is_aws,
#   family = poisson(link = "log"),
#   offset = log(exposure),
#   data = freq_data
# )
# 
# summary(model_freq2)
# 

#ODP check
phi <- sum(residuals(model_freq, type="pearson")^2) / model_freq$df.residual
phi
# phi <- sum(residuals(model_freq2, type="pearson")^2) / model_freq$df.residual
# phi

# Jika phi >= 1, ganti ke Negative Binomial

#Model 1 NB

if (phi > 1.5) {
  model_freq_nb <- glm.nb(
    sum_of_claim_count ~ occupation_code + vehicle_type_code +
      manufacturing_year + excess_type + gender + is_aws +
      offset(log(exposure)),
    data = freq_data
  )
  summary(model_freq_nb)
  use_nb <- TRUE
} else {
  use_nb <- FALSE
}

summary(model_freq_nb)

m_full <- model_freq_nb
m_reduced <- update(m_full, . ~ . - gender)
anova(m_reduced, m_full, test = "Chisq")
# 
# #Model 2 NB
# 
# if (phi > 1.5) {
#   model_freq_nb2 <- glm.nb(
#     sum_of_claim_count ~ vehicle_type_code +
#       manufacturing_year + excess_type + gender + is_aws +
#       offset(log(exposure)),
#     data = freq_data
#   )
#   summary(model_freq_nb2)
#   use_nb <- TRUE
# } else {
#   use_nb <- FALSE
# }
# 
# summary(model_freq_nb2)

tbl_model_freq <- bind_rows(
  cbind(Model="Poisson", model_stats(model_freq)),
  cbind(Model="NegBin",  model_stats(model_freq_nb))
)

write_xlsx(tbl_model_freq, "Tbl Freq.xlsx")

#CHOOSEN MODEL
m_choosen <- model_freq_nb

pred_freq <- read_excel("GLM Claim Data - Frequency.xlsx", sheet = "Prediction Tabel")
freq_factor_cols <- c("occupation_code","vehicle_type_code","excess_type","gender","is_aws")
pred_freq <- align_factor_levels(pred_freq, freq_data, freq_factor_cols)

frequency <- predict(m_choosen, newdata = pred_freq, type = "response")

manufacturing_year <- unique(freq_data$manufacturing_year)


pred_freq_tbl <- data.frame()

for (i in manufacturing_year){
  pred_freq_temp <- pred_freq
  pred_freq_temp[1,3] <- i
  freq_factor_cols <- c("occupation_code","vehicle_type_code","excess_type","gender","is_aws")
  pred_freq <- align_factor_levels(pred_freq_temp, freq_data, freq_factor_cols)
  pred_freq$count <- sum(freq_data$manufacturing_year == i)
  pred_freq$prediction <- predict(m_choosen, newdata = pred_freq_temp, type = "response")
  pred_freq_tbl <- rbind(pred_freq_tbl,pred_freq)
}

write_xlsx(pred_freq_tbl, "Tbl Pred Freq.xlsx")

#SANITY CHECK

phi_nb <- sum(residuals(model_freq_nb, type = "pearson")^2) / df.residual(model_freq_nb)
cat("Dispersion (NegBin):", phi_nb, "\n")


par(mfrow = c(1,2))

#RESIDUAL CHECK
# Pearson residuals vs fitted
plot(fitted(model_freq_nb), residuals(model_freq_nb, type="pearson"),
     xlab="Fitted values", ylab="Pearson residuals",
     main="Residuals vs Fitted (NB)", xlim=c(0,200))  # zoom in
abline(h=0, col="red", lty=2)

# Histogram deviance residuals
hist(residuals(model_freq_nb, type = "deviance"),
     breaks = 30, main="Histogram Deviance Residuals (NB)",
     xlab="Deviance residuals", col="lightblue")

par(mfrow = c(1,1))


#CALIBRATION CHECK
calib_check <- freq_data %>%
  mutate(
    obs  = sum_of_claim_count / exposure,
    pred = fitted(model_freq_nb)
  ) %>%
  group_by(vehicle_type_code) %>%     # bisa diganti grouping lain
  summarise(
    n = n(),
    obs_mean = mean(obs, na.rm=TRUE),
    pred_mean = mean(pred, na.rm=TRUE),
    ratio = pred_mean / pmax(obs_mean, 1e-9),
    .groups="drop"
  )

print(calib_check)

#SEVERITY

model_sev <- glm(
  avg_sev ~ occupation_code + vehicle_type_code +
    manufacturing_year + excess_type + gender + is_aws,
  family = Gamma(link = "log"),
  weights = sum_of_claim_count,   # baris dengan klaim lebih banyak lebih kredibel
  data = sev_data
)

summary(model_sev)

model_sev2 <- glm(
  avg_sev ~ occupation_code + vehicle_type_code +
    manufacturing_year + excess_type + is_aws,
  family = Gamma(link = "log"),
  weights = sum_of_claim_count,   # baris dengan klaim lebih banyak lebih kredibel
  data = sev_data
)

summary(model_sev2)

model_sev3 <- glm(
  avg_sev ~ vehicle_type_code +
    manufacturing_year + excess_type + is_aws,
  family = Gamma(link = "log"),
  weights = sum_of_claim_count,   # baris dengan klaim lebih banyak lebih kredibel
  data = sev_data
)

summary(model_sev3)

model_sev4 <- glm(
  avg_sev ~ vehicle_type_code +
    manufacturing_year  + is_aws,
  family = Gamma(link = "log"),
  weights = sum_of_claim_count,   # baris dengan klaim lebih banyak lebih kredibel
  data = sev_data
)

summary(model_sev4)

tbl_model_sev <- bind_rows(
  cbind(Model="Model Sev 1", model_stats(model_sev)),
  cbind(Model="Model Sev 2", model_stats(model_sev2)),
  cbind(Model="Model Sev 3", model_stats(model_sev3)),
  cbind(Model="Model Sev 4", model_stats(model_sev4))
)

write_xlsx(tbl_model_sev, "Tbl Sev.xlsx")

#CHOOSEN MODEL
m_choosen_sev <- model_sev

pred_sev <- read_excel("GLM Claim Data - Severity.xlsx", sheet = "Prediction Tabel")
sev_factor_cols <- c("occupation_code","vehicle_type_code","excess_type","gender","is_aws")
pred_sev <- align_factor_levels(pred_sev, sev_data, sev_factor_cols)

severity <- predict(m_choosen_sev, newdata = pred_sev, type = "response")

manufacturing_year <- unique(sev_data$manufacturing_year)


pred_sev_tbl <- data.frame()

for (i in manufacturing_year){
  pred_sev_temp <- pred_sev
  pred_sev_temp[1,3] <- i
  sev_factor_cols <- c("occupation_code","vehicle_type_code","excess_type","gender","is_aws")
  pred_sev <- align_factor_levels(pred_sev_temp, sev_data, sev_factor_cols)
  pred_sev$count <- sum(sev_data$manufacturing_year == i)
  pred_sev$prediction <- predict(m_choosen_sev, newdata = pred_sev_temp, type = "response")
  pred_sev_tbl <- rbind(pred_sev_tbl,pred_sev)
}

write_xlsx(pred_sev_tbl, "Tbl Pred Sev.xlsx")

#SANITY CHECK

phi_sev <- sum(residuals(model_sev, type = "pearson")^2) / df.residual(model_sev)
cat("Dispersion (Gamma):", phi_nb, "\n")


par(mfrow = c(1,2))

#RESIDUAL CHECK
# Pearson residuals vs fitted
plot(fitted(model_sev), residuals(model_sev, type="pearson"),
     xlab="Fitted values", ylab="Pearson residuals",
     main="Residuals vs Fitted (NB)")  # zoom in
abline(h=0, col="red", lty=2)

# Histogram deviance residuals
hist(residuals(model_sev, type = "deviance"),
     breaks = 30, main="Histogram Deviance Residuals (NB)",
     xlab="Deviance residuals", col="lightblue")

par(mfrow = c(1,1))


#CALIBRATION CHECK
calib_check <- sev_data %>%
  mutate(
    obs  = avg_sev,
    pred = fitted(model_sev)
  ) %>%
  group_by(vehicle_type_code) %>%     # bisa diganti grouping lain
  summarise(
    n = n(),
    obs_mean = mean(obs, na.rm=TRUE),
    pred_mean = mean(pred, na.rm=TRUE),
    ratio = pred_mean / pmax(obs_mean, 1e-9),
    .groups="drop"
  )

print(calib_check)


pure_premium <- frequency*severity


library(broom)
library(writexl)

write_xlsx(
  list(
    "Frequency Summary" = tidy(model_freq),
    "Severity Summary"  = tidy(model_sev)
  ),
  "glm_summary.xlsx"
)

#ADDITIONAL AT 95% CONFIDENCE INTERVAL

pred_freq <- read_excel("GLM Claim Data - Frequency.xlsx", sheet = "Prediction Tabel")
freq_factor_cols <- c("occupation_code","vehicle_type_code","excess_type","gender","is_aws")
pred_freq <- align_factor_levels(pred_freq, freq_data, freq_factor_cols)

pred_sev <- read_excel("GLM Claim Data - Severity.xlsx", sheet = "Prediction Tabel")
sev_factor_cols <- c("occupation_code","vehicle_type_code","excess_type","gender","is_aws")
pred_sev <- align_factor_levels(pred_sev, sev_data, sev_factor_cols)


z <- 1.96 #95%

# CI untuk FREQUENCY (pakai model m_choosen = NB)
pf <- predict(m_choosen, newdata = pred_freq, type = "link", se.fit = TRUE)
freq_ci <- pred_freq %>%
  mutate(
    freq_hat = exp(pf$fit),
    freq_lo  = exp(pf$fit - z*pf$se.fit),
    freq_hi  = exp(pf$fit + z*pf$se.fit)
  )

write_xlsx(freq_ci, "CI freq.xlsx")

# CI untuk SEVERITY (pakai model m_choosen_sev = Gamma)
ps <- predict(m_choosen_sev, newdata = pred_sev, type = "link", se.fit = TRUE)
sev_ci <- pred_sev %>%
  mutate(
    sev_hat = exp(ps$fit),
    sev_lo  = exp(ps$fit - z*ps$se.fit),
    sev_hi  = exp(ps$fit + z*ps$se.fit)
  )

write_xlsx(sev_ci, "CI sev.xlsx")
