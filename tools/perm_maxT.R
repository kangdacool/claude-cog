##################################
#####  perm_maxT.R  #####
#####  치환 max-T — 상관된 지표 패널에 대한 FWER 통제  #####
##################################
##
## 왜 이 파일이 있는가
## -------------------
## 한 결과변수를 «상관된 지표 여러 개» 에 대해 검정할 때(뇌파 특징 90개, 대사체 200개,
## 유전자 세트 등), BH-FDR 은 독립을 가정해 과잉교정하고 Bonferroni 는 훨씬 더 그렇다.
## 치환 max-T 는 «자료 안의 상관구조를 그대로 써서» FWER 를 정확히 통제한다.
##
## 2026-08-19 한 프로젝트에서 같은 함수를 세 스크립트에 복붙하고 있어서 여기로 뺐다.
##
## 다중검정 «가족» 을 무엇으로 잡을지가 이 함수보다 먼저다 — [[manuscript-rules]] 참조.
## 요약: 가족 = 함께 해석되는 가설의 집합 = 보통 «하나의 결과변수 × 지표 패널».
## 무관한 주제의 검정까지 한 가족에 넣으면 보수적인 게 아니라 비정합이다.
##
## 사용
## ----
##   source("perm_maxT.R")
##   r <- perm_maxT(dat, exposure = "acc_ang", feats = panel90,
##                  covars = c("age", "sex"), nperm = 5000, seed = 1)
##   r$p_fwer   # 패널 전체에 대해 FWER 를 통제한 p
##   r$best     # 최강 지표 이름
##   eff_tests(dat[, panel90])   # 유효 독립검정 수 (Li & Ji) — 함께 보고할 것

##################################
#####  1. 치환 max-T  #####
##################################

#' 한 노출을 지표 패널 전체에 대해 검정하고 FWER 보정 p 를 낸다.
#'
#' 방법: 공변량에 대해 노출·지표를 각각 «잔차화» 한 뒤 노출 잔차만 치환한다
#'       (Freedman-Lane 근사). 통계량 = 패널 전체에서의 max|r|.
#'       관측 max|r| 이상이 나온 치환 비율이 FWER 보정 p 다.
#'
#' @param dat      data.frame
#' @param exposure 노출 열 이름 (1개)
#' @param feats    지표 열 이름 벡터 (패널)
#' @param covars   보정 변수. NULL 이면 무보정
#' @param nperm    치환 횟수 (기본 5000)
#' @param seed     재현용 시드. NULL 이면 설정하지 않음
#' @param min_n    지표당 최소 완전사례 (기본 30)
#' @return list(n, n_feats, feats, obs, best, best_r, p_fwer, null95, crit_r, null_max)
perm_maxT <- function(dat, exposure, feats, covars = NULL,
                      nperm = 5000, seed = NULL, min_n = 30, alpha = 0.05) {
  stopifnot(is.data.frame(dat), length(exposure) == 1L, exposure %in% names(dat))
  if (!is.null(seed)) set.seed(seed)
  covars <- intersect(covars, names(dat))
  feats  <- intersect(feats, names(dat))
  if (!length(feats)) stop("패널에 유효한 지표가 없다")

  keep <- stats::complete.cases(dat[, c(exposure, covars), drop = FALSE])
  d  <- dat[keep, , drop = FALSE]
  if (nrow(d) < min_n) stop("노출·공변량 완전사례가 ", nrow(d), "행뿐이다")

  if (length(covars)) {
    cv <- d[, covars, drop = FALSE]
    x  <- stats::resid(stats::lm(d[[exposure]] ~ ., data = cv))
  } else {
    cv <- NULL
    x  <- d[[exposure]]
  }

  ok <- feats[vapply(feats, function(f) sum(!is.na(d[[f]])) > min_n, TRUE)]
  if (!length(ok)) stop("완전사례가 충분한 지표가 없다")

  Y <- vapply(ok, function(f) {
    y <- d[[f]]
    if (is.null(cv)) return(y)
    m <- stats::complete.cases(y, cv)
    o <- rep(NA_real_, nrow(d))
    o[m] <- stats::resid(stats::lm(y[m] ~ ., data = cv[m, , drop = FALSE]))
    o
  }, numeric(nrow(d)))

  cor_vec <- function(xx) {
    vapply(seq_len(ncol(Y)), function(j) {
      yy <- Y[, j]; m <- !is.na(yy) & !is.na(xx)
      if (sum(m) < min_n) NA_real_ else stats::cor(xx[m], yy[m])
    }, numeric(1))
  }

  obs     <- cor_vec(x)
  obs_max <- max(abs(obs), na.rm = TRUE)
  null_max <- replicate(nperm, max(abs(cor_vec(sample(x))), na.rm = TRUE))

  list(n = nrow(d), n_feats = length(ok), feats = ok, obs = stats::setNames(obs, ok),
       best = ok[which.max(abs(obs))], best_r = obs[which.max(abs(obs))],
       p_fwer = (1 + sum(null_max >= obs_max)) / (nperm + 1),
       null95 = stats::quantile(null_max, 0.95, names = FALSE),
       crit_r = stats::quantile(null_max, 1 - alpha, names = FALSE),
       null_max = null_max)
}

##################################
#####  2. 유효 독립검정 수 (Li & Ji 2005)  #####
##################################
##
## 상관행렬의 고유값에서 «실질» 독립검정 수를 추정한다.
## 치환 p 와 «함께» 보고할 것 — 「통과 N건」 이 실제로 몇 건인지 독자가 알아야 한다.
## 실측 예: 뇌파 특징 명목 90 → 실질 30, 그중 연결성 24 → 실질 6.

eff_tests <- function(M) {
  M <- as.data.frame(M)
  keep <- vapply(M, function(x) is.numeric(x) && stats::sd(x, na.rm = TRUE) > 0, TRUE)
  M <- M[, keep, drop = FALSE]
  if (!ncol(M)) return(0)
  R <- stats::cor(M, use = "pairwise.complete.obs")
  R[is.na(R)] <- 0
  ev <- pmax(eigen(R, symmetric = TRUE, only.values = TRUE)$values, 0)
  sum(ifelse(ev >= 1, 1, 0) + (ev - floor(ev)))
}

##################################
#####  3. 여러 가족을 나란히 놓고 비교  #####
##################################
##
## 「이 q 값이 맞아?」 라는 질문이 나오면 답은 «가족을 무엇으로 잡았나» 다.
## 후보를 나란히 계산해 근거를 갖고 고르라. 결과를 보고 고르지 말 것 —
## 후보 목록과 채택 기준을 «먼저» 적고 실행한다.
##
#' @param p       p값 벡터
#' @param families named list. 각 원소는 p 와 같은 길이의 «가족 라벨» 벡터
#'                (예: list(global = rep("all", n), by_outcome = outcome_vec))
#' @return data.frame(family, n_tests, n_sig_05, min_q)
compare_families <- function(p, families, alpha = 0.05) {
  do.call(rbind, lapply(names(families), function(nm) {
    lab <- families[[nm]]
    q <- rep(NA_real_, length(p))
    for (k in unique(lab[!is.na(lab)])) {
      i <- !is.na(lab) & lab == k & !is.na(p)
      if (any(i)) q[i] <- stats::p.adjust(p[i], method = "BH")
    }
    data.frame(family = nm, n_groups = length(unique(lab[!is.na(lab)])),
               n_tests = sum(!is.na(p)), n_sig = sum(q < alpha, na.rm = TRUE),
               min_q = suppressWarnings(min(q, na.rm = TRUE)),
               stringsAsFactors = FALSE)
  }))
}

cat("--- perm_maxT.R loaded (perm_maxT / eff_tests / compare_families)\n")
