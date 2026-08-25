## ============================================================================
## sm_exposure_diagnostic.R — 통계적 매칭(SM) 노출 후보 사전진단 (재사용 도구)
##
## 2026-08-19 개발. 근거·한계는 [[sm-analysis-stage-pitfalls]] ⑤ 참조.
##
## 왜 필요한가
##   CIA 하에서 융합 연관성 = cov(E[Z|X], E[Y|X]) — 셀 변수 X의 기울기뿐이다.
##   노출 Z가 인구학 X의 함수이면(=인구학이 잘 설명하면) 나오는 결과는 인구학의 그림자다.
##   2026-08-19에 이 진단 없이 4연속 실패했다(의료신뢰→고혈압 / 혼자시간→외로움 /
##   디지털역량→혈압 / 직장차별→대사지표). 전부 착수 전 이 함수 하나로 걸러졌을 것이다.
##
## ⚠ 한계 — 반드시 알고 쓸 것
##   이 지표는 "비인구학 지역변이가 있는가"를 잰다. **성공을 예측하지는 못한다**
##   (결과 쪽 지역변이와의 상관이 결정적이고 그건 붙여보기 전엔 모름).
##   실제로 한 분석에서 성공한 노출은 CHS 102개 중 68~87위였다(절대값은 양호:
##   rel 0.90~0.96, R2geo_c 0.33~0.58). ⇒ **후보 풀 생성용으로만 쓰고 순위를 맹신하지 말 것.**
##
## ⚠ 개인수준 R²를 쓰지 말 것 (초판이 그래서 자기검증에 실패했다)
##   개인 분산이 압도적이라 개인 R²_geo는 어떤 변수든 0.01 수준으로 나온다.
##   셀 기반 매칭이 쓰는 것은 **셀 평균**이므로 표본오차를 제거한 셀평균 분산으로 재야 한다.
##
## 사용 예
##   source("sm_exposure_diagnostic.R")
##   out <- sm_diag(d, exposures = cand_vars, geo = "signgu_code",
##                  age = "age", sex = "sex", edu = "sob_01z1",
##                  miss_codes = c(7,8,9,77,88,99),
##                  known_good = c("ord_01d2","ord_01f3"))   # 검증용 known-good 사례
## ============================================================================

sm_diag <- function(data, exposures, geo, age, sex, edu,
                    miss_codes = c(7, 8, 9, 77, 88, 99, 777, 888, 999),
                    min_n = 5000, min_cell_n = 100, min_cells = 20,
                    known_good = NULL, verbose = TRUE) {
  stopifnot(all(c(geo, age, sex, edu) %in% names(data)))
  num <- function(x) suppressWarnings(as.numeric(x))

  base <- data.frame(
    age = num(data[[age]]),
    sex = num(data[[sex]]),
    edu = num(data[[edu]]),
    geo = as.character(data[[geo]])
  )
  base$edu[base$edu %in% miss_codes | base$edu > 90] <- NA
  ok_base <- stats::complete.cases(base)
  if (verbose) cat("[sm_diag] 기준변수 완전관측:", sum(ok_base),
                   "| 셀 수:", length(unique(base$geo[ok_base])), "\n")

  res <- vector("list", length(exposures))
  for (i in seq_along(exposures)) {
    v <- exposures[i]
    if (!v %in% names(data)) next
    x <- num(data[[v]]); x[x %in% miss_codes] <- NA
    ok <- ok_base & !is.na(x)
    if (sum(ok) < min_n) next
    z <- x[ok]; b <- base[ok, ]
    if (length(unique(z)) < 2) next

    cl <- stats::aggregate(
      cbind(z = z, age = b$age, fem = as.integer(b$sex == 2), edu = b$edu) ~ geo,
      data = data.frame(z = z, age = b$age, sex = b$sex, edu = b$edu, geo = b$geo),
      FUN = mean)
    cnt <- as.data.frame(table(b$geo), stringsAsFactors = FALSE)
    names(cnt) <- c("geo", "n")
    vr  <- stats::aggregate(z ~ geo, data = data.frame(z = z, geo = b$geo), FUN = stats::var)
    names(vr)[2] <- "v"
    cl <- merge(merge(cl, cnt, by = "geo"), vr, by = "geo")
    cl <- cl[cl$n >= min_cell_n, ]
    if (nrow(cl) < min_cells) next

    nbar   <- mean(cl$n)
    var_w  <- mean(cl$v, na.rm = TRUE)
    var_b  <- stats::var(cl$z)
    var_bt <- max(var_b - var_w / nbar, 0)              # 표본오차 제거한 진짜 between
    rel    <- var_bt / (var_bt + var_w / nbar)          # 셀평균 신뢰도
    icc    <- var_bt / (var_bt + var_w)

    m_c <- stats::lm(z ~ age + fem + edu, data = cl)    # 셀의 인구학 구성으로 설명
    r2demo <- summary(m_c)$r.squared
    var_rt <- max(stats::var(stats::resid(m_c)) - var_w / nbar, 0)
    r2geo  <- var_rt / (var_bt + 1e-12)                 # 인구학 통제 후 남는 진짜 지역변이

    ## 노출군-비노출군 연령차(|5세| 넘으면 위험 신호)
    if (length(unique(z)) == 2) {
      lv <- sort(unique(z)); gap <- mean(b$age[z == lv[2]]) - mean(b$age[z == lv[1]])
    } else {
      md <- stats::median(z); gap <- mean(b$age[z > md]) - mean(b$age[z <= md])
    }

    ## ---- 관문4(노출 밀도) 지표 : 2026-08-19 추가 ----
    ## rel은 곧 **회귀계수 감쇠계수**다(고전적 측정오차). 관측 기울기 ≈ rel x 진짜 기울기.
    ## 희소 노출(0이 다수)은 var_w가 커서 rel이 떨어지고, 진짜 효과도 그 비율로 깎인다.
    zero_frac <- mean(z == min(z, na.rm = TRUE))          # 최빈 하한값(대개 0) 비율
    ## rel=0.8을 맞추려면 셀당 몇 명이 필요한가  (rel = vb/(vb+vw/n) = .8  ->  n = 4*vw/vb)
    n_need <- if (var_bt > 0) ceiling(4 * var_w / var_bt) else NA_integer_

    lb <- attr(data[[v]], "label"); lb <- if (is.null(lb)) "" else as.character(lb)
    res[[i]] <- data.frame(var = v, n = sum(ok), n_cell = nrow(cl), n_per_cell = round(nbar),
                           zero_frac = zero_frac,
                           ICC = icc, rel = rel, atten = rel, n_need80 = n_need,
                           R2demo_c = r2demo, R2geo_c = r2geo,
                           age_gap = gap, score = rel * r2geo,
                           label = substr(lb, 1, 50), stringsAsFactors = FALSE)
  }
  out <- do.call(rbind, res)
  if (is.null(out)) { warning("진단 가능한 후보 없음"); return(invisible(NULL)) }
  out <- out[order(-out$score), ]
  out$rank <- seq_len(nrow(out))

  if (!is.null(known_good) && verbose) {
    cat("\n[sm_diag] ⭐ known-good 검증 (지표가 타당하면 상위권이어야 함):\n")
    kg <- out[out$var %in% known_good,
              c("rank","var","rel","zero_frac","n_per_cell","R2demo_c","R2geo_c","score","age_gap","label")]
    if (nrow(kg) == 0) cat("  known_good 변수가 결과에 없음\n") else print(kg, row.names = FALSE)
    cat("  (전체", nrow(out), "개 중 순위. 하위권이면 지표를 의심할 것 —",
        "실제로 초판은 이 검증으로 폐기됐다)\n")
    ## ⭐ 기각 감사: known-good을 죽이는 문턱은 그 자체로 무효다
    if (nrow(kg) > 0) {
      cat("\n[sm_diag] ⚠ 기각 문턱 자기검사 — 아래 값보다 엄격한 컷은 known-good도 죽인다:\n")
      cat(sprintf("    rel >= %.2f | R2geo_c >= %.2f | |age_gap| <= %.1f | zero_frac <= %.2f\n",
                  min(kg$rel), min(kg$R2geo_c), max(abs(kg$age_gap)), max(kg$zero_frac)))
      cat("    후보를 죽이려 할 때 쓰는 기준이 이 줄을 넘으면 **기준이 틀린 것이다.**\n")
    }
  }
  out
}

## 해석 가이드
##  R2demo_c 높음(>0.3) + |age_gap|>5  -> 인구학의 그림자. **버릴 것.**
##  R2geo_c 높음(>0.5) + rel>0.8       -> 비인구학 지역변이 확보. 후보 풀에 넣을 것.
##  rel 낮음(<0.5)                      -> 셀 평균이 표본오차에 묻힘. 셀을 키우거나 포기.
##
## ⭐ 관문4 (노출 밀도) — 2026-08-19 추가. 근거: [[us-exposure-density-bind]]
##  `atten`(=rel)은 **회귀계수 감쇠계수**다. 관측 기울기 ≈ atten x 진짜 기울기.
##  착수 전에 이걸 보면 "귀무가 나올 것"을 미리 안다. 실측 대조(같은 파이프라인·같은 셀구조):
##    ATUS TV시간   zero_frac 0.22, 셀당 64 -> 비만 p=3.9e-06 (살았다)
##    ATUS 통근시간 zero_frac 0.78, 셀당 49 -> 전부 귀무    (감쇠로 죽었다)
##  `n_need80`은 rel=0.8을 맞추는 데 필요한 셀당 표본수. 현재 `n_per_cell`이 그보다 훨씬
##  작으면 **셀을 키우거나(연령대 병합·지리 상위단위) 자료를 늘리기 전엔 돌리지 말 것.**
##  ⚠ 단 이건 "죽었다"가 아니라 "검정력 부족"이다 — 귀무를 근거로 주제를 폐기하지 말 것.
