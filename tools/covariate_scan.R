#!/usr/bin/env Rscript
##################################################
#####  covariate_scan.R — Table 1 표준 축 점검  #####
##################################################
# 무엇을 하는가
#   자료 파일 하나를 받아 세 가지를 본다.
#     §3 축 대조    — 역학 Table 1 표준 축이 있나 없나. 절대 잘라서 출력하지 않는다.
#     §4 척도 정합  — 열 «이름»이 주장하는 척도를 «값»이 배신하는가.
#     §5 개념 중복  — 한 축에 열이 여럿일 때 그것들이 서로 같은가.
#
# 왜 필요한가 (세 가지가 각각 실제 사고에서 나왔다)
#   §3  공변량 목록을 「이 분석에 뭐가 필요한가」로 만들면 인구사회 기본 변수를 빠뜨린다.
#       Table 1 은 분석이 아니라 «자료»가 정한다.
#       2026-08-21 KTLS: 최종학력·혼인·자녀·월소득이 네 파고에 다 있었는데 못 찾고 모형을
#       성별·담임·학급규모만으로 돌리고 있었다. 원인 둘 -- (1) 진단 출력을 head 로 잘라
#       읽었다 (2) 애초에 소득·혼인을 검색어에 넣지 않았다.
#   §4  ⭐ 「열이 있다」를 확인하고 «그 열이 무엇인지»는 이름을 믿는다 -- 이게 반복된다.
#       2026-08-22: `SNSB_edu_2y` 를 교육«연수»로 써서 프로젝트 두 개가 1~7 범주코드로
#       모형을 돌렸다(진짜 연수는 `eduyr_snsb_2y`, 평균 10.8). 유의성 5건이 뒤집혔다.
#       2026-08-23: `_s_00` 을 「규준 z」로 소개했는데 14열 전부 [0.01, 99.99]·평균 50 인
#       백분위였다. **숫자 추적 검사는 이걸 못 잡는다** -- 값은 표에 실재하고 «이름»이 틀렸다.
#   §5  한 개념에 열이 여럿이면 어느 것이 정본인지 자료가 말해주지 않는다.
#       2026-08-23: MCI 판정 열이 회차마다 3~4개였다. `mci_2y` 와 `mci_more_2y` 는
#       완전히 같은 열인데 `mci_gds_2y` 는 **4명 중 1명에서 갈리고** 유병이 58.6% ↔ 51.4% 다.
#       사전지정 없이 하나를 고르면 그게 결과를 정한다.
#
# 사용법
#   Rscript covariate_scan.R <file.sav|.sas7bdat|.rds|.csv> [--extra <정규식>] [--axes-only]
#   --axes-only  §4·§5 를 건너뛴다(전체 읽기를 피한다 -- 아주 큰 파일에서만)
#
# ⚠ 이 스크립트는 «후보»를 준다. 판정은 사람이 한다 -- 정규식은 축을 맞히지 못할 수 있고,
#   정상적으로 작은 값을 갖는 변수도 §4 의 신호를 낸다.
##################################################

args <- commandArgs(trailingOnly = TRUE)
if (!length(args)) {
  cat("사용법: Rscript covariate_scan.R <file.sav|.sas7bdat|.rds|.csv> [--extra <정규식>] [--axes-only]\n")
  quit(status = 1)
}
path <- args[1]
extra <- if ("--extra" %in% args) args[which(args == "--extra") + 1] else NULL
axes_only <- "--axes-only" %in% args

if (!file.exists(path)) stop("파일이 없다: ", path)

##################################
#####  1. 표준 축  #####
##################################
# 역학 논문 Table 1 이 거의 언제나 담는 것들. 프로젝트마다 더할 수는 있어도 뺄 수는 없다.
AXES <- list(
  "성별"            = "성별|남녀|sex|gender",
  "연령"            = "연령|나이|출생|(^|[^a-z])ages?($|[^a-z])",
  # ⚠ `age` 를 맨몸으로 쓰면 **average** 에 걸린다 -- 2026-08-23 한 프로젝트에서 관절각도
  #   라벨("average ...")이 2,460 열이나 「연령」으로 잡혔다. 앞뒤를 비문자로 묶는다.
  "학력"            = "학력|학위|교육수준|졸업|education|edu",
  "소득"            = "소득|급여|연봉|보수|임금|월급|income|earning|salary|wage",
  "혼인/가구"       = "혼인|결혼|배우자|이혼|사별|marit|marri|spouse",
  "자녀/부양"       = "자녀|아동수|부양|child",
  "고용형태"        = "고용형태|정규직|비정규|계약|임용형태|employ",
  "근로시간"        = "근로시간|근무시간|주당.*시간|초과근무|working hour|worktime|wtime",
  "직업/직종"       = "직종|직업|occupation|occ|job",
  "지역"            = "지역|시도|거주지|도시|region|urban",
  "건강행태(흡연)"  = "흡연|담배|smok",
  "건강행태(음주)"  = "음주|술|alcohol|drink",
  "신체활동"        = "운동|신체활동|physical activity|exercise",
  "주관적 건강"     = "주관적 건강|건강수준|self-rated health|srh",
  "만성질환"        = "만성|진단|고혈압|당뇨|질환|chronic|disease"
)
if (!is.null(extra)) AXES[["(사용자 지정)"]] <- extra

##################################
#####  2. 자료 읽기  #####
##################################
# §4·§5 는 «값»을 봐야 하므로 전체를 읽는다. 라벨만 필요하면 --axes-only 로 피한다.
ext <- tolower(tools::file_ext(path))
if (ext %in% c("sav", "sas7bdat")) {
  if (!requireNamespace("haven", quietly = TRUE)) stop("haven 패키지가 필요하다")
  d <- if (ext == "sav") haven::read_sav(path, n_max = if (axes_only) 5L else Inf) else
       haven::read_sas(path)
  if (axes_only && nrow(d) > 5) d <- d[1:5, , drop = FALSE]
} else if (ext == "rds") {
  d <- readRDS(path)
  if (axes_only && nrow(d) > 5) d <- d[1:5, , drop = FALSE]
} else if (ext == "csv") {
  d <- utils::read.csv(path, nrows = if (axes_only) 5 else -1, fileEncoding = "UTF-8-BOM")
} else {
  stop("지원하지 않는 확장자: ", ext, " (sav/sas7bdat/rds/csv)")
}

lab <- vapply(d, function(x) {
  a <- attr(x, "label")
  if (is.null(a)) "" else as.character(a)[1]
}, character(1))
# 라벨이 없는 자료(csv 등)에서는 변수명 자체를 검색 대상으로 삼는다.
hay <- ifelse(nzchar(lab), paste(names(lab), lab), names(lab))

cat(sprintf("\n파일: %s\n행 %s · 변수 %d개 · 라벨 있는 변수 %d개\n",
            basename(path), if (axes_only) "(축 점검만)" else format(nrow(d), big.mark = ","),
            length(hay), sum(nzchar(lab))))
cat(strrep("=", 78), "\n", sep = "")

##################################
#####  3. 축별 대조 (자르지 않는다)  #####
##################################
axis_hits <- list()
missing_axes <- character(0)
for (ax in names(AXES)) {
  idx <- grep(AXES[[ax]], hay, perl = TRUE, ignore.case = TRUE)
  if (!length(idx)) {
    cat(sprintf("\n[없음] %s\n", ax))
    missing_axes <- c(missing_axes, ax)
    next
  }
  axis_hits[[ax]] <- names(lab)[idx]
  cat(sprintf("\n[있음] %s  (%d개)\n", ax, length(idx)))
  for (i in idx) {
    txt <- if (nzchar(lab[[i]])) lab[[i]] else "(라벨 없음)"
    cat(sprintf("   %-14s %s\n", names(lab)[i], substr(txt, 1, 84)))
  }
}

cat("\n", strrep("=", 78), "\n", sep = "")
if (length(missing_axes)) {
  cat("자료에 «없는» 축:\n  ", paste(missing_axes, collapse = " · "), "\n", sep = "")
  cat("\n이 목록은 «자료에 없다»는 뜻이지 «찾다 말았다»가 아니다.\n")
  cat("Table 1 이나 보정에서 빠뜨릴 때는 여기에 근거를 대고 빠뜨릴 것.\n")
} else {
  cat("표준 축이 모두 존재한다.\n")
}
cat("\n⚠ 걸린 변수가 «그 축이 맞는지»는 사람이 본다. 정규식은 후보를 줄 뿐이다.\n")

if (axes_only) {
  cat("\n(--axes-only: 척도 정합·개념 중복 점검을 건너뛰었다)\n")
  quit(status = 0)
}

num_of <- function(x) suppressWarnings(as.numeric(if (is.factor(x)) as.character(x) else x))

##################################
#####  4. 척도 정합 — 이름이 주장하는 것을 값이 배신하는가  #####
##################################
# 이름이 척도를 «주장»하면 값이 그 주장과 맞아야 한다. 안 맞는 것만 찍는다.
# (전 열을 찍으면 아무도 안 읽는다 -- profile_columns.py 의 --only-suspect 와 같은 이유.)
CLAIMS <- list(
  list(nm = "z 점수",
       re = "(^|_)z($|_)|zscore|z_score|std_?score|표준점수|표준화",
       ok  = function(mn, mx, v) max(abs(c(mn, mx))) <= 10,
       msg = "z 라면 대개 |값| <= 6 이다"),
  list(nm = "백분위",
       re = "백분위|percentile|pctl",
       ok  = function(mn, mx, v) mn >= -0.5 && mx <= 100.5,
       msg = "백분위라면 0~100 이어야 한다"),
  list(nm = "퍼센트",
       re = "percent|퍼센트|(^|_)pct($|_)",
       ok  = function(mn, mx, v) mn >= -0.5 && mx <= 100.5,
       msg = "퍼센트라면 0~100 이어야 한다"),
  list(nm = "비율(0~1)",
       # 로그·제곱근 변환본은 애초에 0~1 이 아니다 -- 이름에 그렇게 적혀 있으면 봐준다.
       re = "((^|_)(prop|ratio|frac)($|_)|비율)(?!.*(log|sqrt|_ln($|_)))",
       ok  = function(mn, mx, v) mx <= 1.5,
       msg = "0~1 비율로 읽히는 이름인데 값이 더 크다 -- 퍼센트인가"),
  list(nm = "연수(년)",
       re = "(^|_)(yr|yrs|year|years)($|_)|연수|년수",
       ok  = function(mn, mx, v) !(mx <= 7 && length(unique(v)) >= 3 &&
                                   all(abs(v - round(v)) < 1e-8)),
       msg = "「연수」인데 작은 정수 코드로 보인다 -- 범주코드가 아닌가")
)

cat("\n", strrep("=", 78), "\n", sep = "")
cat("§4 척도 정합 — 이름이 주장하는 척도를 값이 배신하는 열\n\n")
flagged <- 0L
for (j in seq_along(d)) {
  v <- num_of(d[[j]])
  v <- v[is.finite(v)]
  if (length(v) < 10 || length(unique(v)) < 2) next
  nmj <- names(d)[j]
  target <- if (nzchar(lab[[j]])) paste(nmj, lab[[j]]) else nmj
  mn <- min(v); mx <- max(v)
  for (cl in CLAIMS) {
    if (!grepl(cl$re, target, perl = TRUE, ignore.case = TRUE)) next
    if (isTRUE(cl$ok(mn, mx, v))) next
    flagged <- flagged + 1L
    cat(sprintf("  [%s] %-28s n=%5d  [%.2f, %.2f]  평균 %.2f\n",
                cl$nm, substr(nmj, 1, 28), length(v), mn, mx, mean(v)))
    cat(sprintf("        -> %s\n", cl$msg))
  }
}
if (!flagged) cat("  이름과 값이 어긋나는 열은 없다.\n")
cat(sprintf("\n  검사한 주장 유형 %d 종 · 표시된 열 %d 개\n", length(CLAIMS), flagged))
cat("  ⚠ 정상인데 걸릴 수 있다(예: 0~1 만 관측된 백분위). 반대로 이름이 척도를\n")
cat("     주장하지 «않으면» 이 검사는 침묵한다 -- 침묵이 무죄는 아니다.\n")

##################################
#####  5. 개념 중복 — 한 축에 열이 여럿일 때 서로 같은가  #####
##################################
# 「있다」와 「하나다」는 다르다. 같은 축의 열들이 서로 다른 사람을 가리키면
# 어느 것을 쓰느냐가 결과를 정한다 -- 그 선택은 «분석 전에» 사전지정돼야 한다.
cat("\n", strrep("=", 78), "\n", sep = "")
cat("§5 개념 중복 — 같은 축에 열이 여럿인 경우\n\n")
MISMATCH_N <- 30L  # 이보다 많으면 축 매칭 실패로 본다
MAXCOL  <- 12L   # 축 하나에 열이 수십 개면 나열이 무의미하다
MAXPAIR <- 8L
any_dup <- FALSE
for (ax in names(axis_hits)) {
  cols <- Filter(function(cn) {
    v <- num_of(d[[cn]]); sum(is.finite(v)) >= 10 && length(unique(v[is.finite(v)])) >= 2
  }, axis_hits[[ax]])
  if (length(cols) < 2) next
  any_dup <- TRUE
  # 축 하나가 수십 열이면 「개념이 중복」된 게 아니라 «정규식이 축을 잘못 잡은» 것이다.
  # 그 경우 쌍 대조는 의미가 없고 출력만 덮는다 -- 그렇게 말하고 넘어간다.
  if (length(cols) > MISMATCH_N) {
    cat(sprintf("[%s]  수치형 열 %d 개 -- ⚠ 너무 많다. 정규식이 이 축을 잘못 잡았을 수 있다\n",
                ax, length(cols)))
    cat(sprintf("        (앞 3 개: %s ...)\n\n", paste(utils::head(cols, 3), collapse = ", ")))
    next
  }
  cat(sprintf("[%s]  수치형 열 %d 개%s\n", ax, length(cols),
              if (length(cols) > MAXCOL) sprintf(" (앞 %d 개만 보인다)", MAXCOL) else ""))
  shown <- utils::head(cols, MAXCOL)
  for (cn in shown) {
    v <- num_of(d[[cn]]); v <- v[is.finite(v)]
    cat(sprintf("   %-26s n=%5d  [%8.2f, %8.2f]  평균 %8.2f  고유값 %d\n",
                substr(cn, 1, 26), length(v), min(v), max(v), mean(v), length(unique(v))))
  }
  # 쌍별 대조. 이진이면 «불일치 인원», 연속이면 상관 + 값이 같은 비율.
  pr <- utils::combn(shown, 2, simplify = FALSE)
  if (length(pr) > MAXPAIR) {
    cat(sprintf("   (쌍 %d 개 중 앞 %d 개만 대조한다 -- 나머지는 사람이 고를 것)\n",
                length(pr), MAXPAIR))
    pr <- pr[seq_len(MAXPAIR)]
  }
  for (p in pr) {
    a <- num_of(d[[p[1]]]); b <- num_of(d[[p[2]]])
    k <- is.finite(a) & is.finite(b)
    if (sum(k) < 10) next
    if (all(a[k] %in% c(0, 1)) && all(b[k] %in% c(0, 1))) {
      cat(sprintf("   %-22s vs %-22s 공통 %4d · 불일치 %3d (%4.1f%%)\n",
                  substr(p[1], 1, 22), substr(p[2], 1, 22), sum(k),
                  sum(a[k] != b[k]), 100 * mean(a[k] != b[k])))
    } else {
      # 한쪽이 상수면 cor 은 NA 와 경고를 낸다 -- 경고 대신 «상수»라고 말한다.
      rr <- if (stats::sd(a[k]) == 0 || stats::sd(b[k]) == 0) NA_real_ else stats::cor(a[k], b[k])
      cat(sprintf("   %-22s vs %-22s 공통 %4d · r=%s · 값이 같은 비율 %5.1f%%\n",
                  substr(p[1], 1, 22), substr(p[2], 1, 22), sum(k),
                  if (is.na(rr)) " 상수 " else sprintf("%6.3f", rr),
                  100 * mean(a[k] == b[k])))
    }
  }
  cat("\n")
}
if (!any_dup) cat("  한 축에 수치형 열이 둘 이상인 경우가 없다.\n")
cat("⭐ 불일치가 크면 «어느 정의를 primary 로 쓰는가»를 분석 전에 정하고 사전명세에 적는다.\n")
cat("   불일치 0 인 쌍은 같은 열의 다른 이름일 수 있다 -- 둘 다 보고할 필요가 없다.\n")
