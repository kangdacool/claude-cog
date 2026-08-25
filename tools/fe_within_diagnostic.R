##################################
#####  FE_WITHIN_DIAGNOSTIC  #####
##################################
# 개인 고정효과(FE) 추정치를 «손으로 만든 진단»으로 설명할 때 쓰는 헬퍼.
#
# 🔴 왜 있나 (2026-08-25 실사고):
#   `feols(y ~ x | id + wave)` 의 계수가 «왜» 그 값인지 보려고 손으로 개인내 상관·
#   VIF·부분 감쇠를 계산했는데, **사람 평균만 빼고 차수 평균을 안 뺐다.** 결과가
#   수직척도(학년마다 오르는 검사점수)라 공통 추세가 남아 개인내 상관이 +0.005,
#   부분 계수가 «양수»로 나왔다. 그대로 보고했으면 **방향을 정반대로** 말할 뻔했다.
#
#   교훈은 「양방향 차분을 기억하라」가 아니다 -- 기억은 매 세션 리셋된다.
#   ⭐ **손으로 만든 진단은 «그 모형의 계수를 재현하는지» 먼저 확인해야 한다.**
#      재현하지 못하면 그 진단은 다른 것을 설명하고 있다. 그 확인을 코드로 박는다.
#
# USAGE
#   source("<repo>/agent/tools/fe_within_diagnostic.R")
#   W <- fe_within(dat, y = "z_achv", x = "z_lonely", id = "sid", time = "wave",
#                  covs = c("x_bully", "x_alien"), cluster = "school")
#   W$data      # 양방향 차분된 프레임 (여기서 상관·VIF·부분감쇠를 본다)
#   W$check     # 재현 검사 결과 (모형 계수 vs 차분 OLS 계수)
##################################

fe_within <- function(dat, y, x, id, time, covs = character(0),
                      cluster = NULL, tol = 1e-6, verbose = TRUE) {
  need <- c(y, x, id, time, covs, cluster)
  miss <- setdiff(need, names(dat))
  if (length(miss)) stop("열이 없다: ", paste(miss, collapse = ", "))
  d <- dat[stats::complete.cases(dat[, need, drop = FALSE]), , drop = FALSE]

  # 🔴 singleton 제거. `fixest` 는 고정효과 수준이 관측 1개뿐인 사람을 «버린다».
  #    손 차분이 그들을 남기면 계수가 달라진다(2026-08-25 실측: feols −0.3508 vs 손 −0.3232).
  #    반복해서 제거한다 -- 하나를 빼면 다른 수준이 새로 singleton 이 될 수 있다.
  repeat {
    n0 <- nrow(d)
    d <- d[d[[id]] %in% names(which(table(d[[id]]) >= 2)), , drop = FALSE]
    d <- d[d[[time]] %in% names(which(table(d[[time]]) >= 2)), , drop = FALSE]
    if (nrow(d) == n0 || nrow(d) == 0) break
  }
  if (nrow(d) == 0) stop("singleton 제거 후 남은 관측이 없다")

  # 🔴 «교대 차분»(alternating projections). 불균형 패널에서 사람->차수 한 번씩만 빼면
  #    정확한 within 변환이 아니다 -- 수렴할 때까지 번갈아 뺀다.
  dm2 <- function(v) {
    r <- v
    for (i in 1:200) {
      prev <- r
      r <- r - ave(r, d[[id]], FUN = mean)
      r <- r - ave(r, d[[time]], FUN = mean)
      if (max(abs(r - prev)) < 1e-10) break
    }
    r
  }
  W <- data.frame(setNames(lapply(c(y, x, covs), function(v) dm2(d[[v]])), c(y, x, covs)))
  W[[id]] <- d[[id]]; W[[time]] <- d[[time]]
  if (!is.null(cluster)) W[[cluster]] <- d[[cluster]]

  # ---- 재현 검사: 차분 OLS 가 feols 의 FE 계수를 되살리는가 ----
  chk <- list(ok = NA, model_est = NA_real_, demeaned_est = NA_real_)
  rhs <- paste(c(x, covs), collapse = " + ")
  b_dm <- unname(coef(stats::lm(stats::as.formula(sprintf("%s ~ %s", y, rhs)), data = W))[x])
  chk$demeaned_est <- b_dm
  if (requireNamespace("fixest", quietly = TRUE)) {
    f <- stats::as.formula(sprintf("%s ~ %s | %s + %s", y, rhs, id, time))
    m <- try(fixest::feols(f, data = d), silent = TRUE)
    if (!inherits(m, "try-error")) {
      b_fe <- unname(coef(m)[x])
      chk$model_est <- b_fe
      chk$ok <- isTRUE(abs(b_fe - b_dm) < max(tol, 1e-3 * abs(b_fe)))
    }
  }
  if (isFALSE(chk$ok)) {
    stop(sprintf(paste("🔴 진단이 모형을 재현하지 못한다: feols %.6f vs 차분 OLS %.6f.",
                       "\n   이 차분 프레임으로 계산한 상관·VIF·부분감쇠는 «그 모형의」 것이",
                       "아니다. 불균형이 심하거나 고정효과 구성이 다르다."),
                 chk$model_est, chk$demeaned_est))
  }
  if (verbose && isTRUE(chk$ok)) {
    cat(sprintf("  ✅ 재현 검사 통과: feols %.6f == 차분 OLS %.6f\n",
                chk$model_est, chk$demeaned_est))
  }

  # ---- 노출의 식별 변동이 공변량에 얼마나 먹히는가 (과보정 vs 공선성 판별) ----
  absorbed <- NA_real_; vif_x <- NA_real_
  if (length(covs)) {
    r2 <- summary(stats::lm(stats::as.formula(
      sprintf("%s ~ %s", x, paste(covs, collapse = " + "))), data = W))$r.squared
    absorbed <- r2; vif_x <- 1 / (1 - r2)
    if (verbose) {
      cat(sprintf("  노출의 개인내 변동 중 공변량이 먹는 비율 R2=%.3f (VIF %.2f) -> 잔존 %.1f%%\n",
                  r2, vif_x, 100 * (1 - r2)))
      cat("  ⚠ VIF가 낮은데 계수가 크게 줄면 공선성이 아니라 «과보정/구성개념 중첩»을 의심한다.\n")
    }
  }
  list(data = W, check = chk, exposure_absorbed_r2 = absorbed, exposure_vif = vif_x, n = nrow(W))
}
