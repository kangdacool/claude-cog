# ==============================================================================
# r_longrun.R -- 며칠짜리 무인 R 실행이 끊겨도 살아남게 하는 최소 헬퍼 4개
#
# WHY. 2026-08-17~19, 38셀 x 200 replicate짜리 시뮬레이션이 28.5시간 만에
# exit 1로 죽었다. 로그에 에러 한 줄 없었다. 되살리는 과정에서 네 가지가 물렸고, 넷 다
# «짤 때 정해지는» 것이라 나중에 고칠 수 없었다. 그 넷을 여기 박아 둔다.
#
#   1. 캐시 조회를 오늘 날짜로만 하면, 자정을 넘긴 작업은 다음 날 재시작할 때 어제 끝낸
#      것을 못 찾는다. (14셀 ~30시간을 조용히 재계산할 뻔했다.)
#   2. parLapply는 워커 하나가 죽으면 마스터까지 내려가고 수집분도 같이 날아간다.
#   3. 작업 단위 캐시만 있으면 «작업 안에서» 죽었을 때 그 몇 시간이 통째로 날아간다.
#   4. 완료 시 원시 draw를 지우면, 나중에 필요한 분석(짝 대비 등)이 영영 불가능해진다.
#
# 이 파일은 의존성이 없다(base + parallel). `source()` 해서 쓴다.
#
# 셀프테스트:  Rscript r_longrun.R --selftest
# ==============================================================================


# ---- 1. 날짜 무관 캐시 -------------------------------------------------------
# 쓸 때는 오늘 스탬프, 읽을 때는 아무 스탬프나.
#
# ⚠ 고정문자열 접두 매칭을 쓴다. 태그를 정규식에 끼워 넣으려다 이스케이프를 잘못 만들어
#   "invalid regular expression"으로 죽은 적 있다(태그에 "." 이 들어가는 경우 — 분수 델타 등).

ck_path <- function(dir, tag, stamp = format(Sys.Date(), "%y%m%d"), ext = "rds") {
  file.path(dir, sprintf("%s_%s.%s", tag, stamp, ext))
}

#' 완료 산출물을 «날짜와 무관하게» 찾는다. 여러 개면 가장 최근 것.
#' @return 경로 문자열, 없으면 NULL
ck_find <- function(dir, tag, ext = "rds") {
  pref <- paste0(tag, "_")
  suf  <- paste0(".", ext)
  cand <- list.files(dir)
  hit  <- cand[startsWith(cand, pref) & endsWith(cand, suf)]
  if (!length(hit)) return(NULL)
  full <- file.path(dir, hit)
  full[which.max(file.mtime(full))]
}

#' 캐시가 있으면 읽고, 없으면 계산해서 저장한다.
#' @param save_raw  선택. 원시 중간자료(리스트). 주면 `{tag}_raw_{stamp}` 로 따로 저장한다.
#'                  «요약만 있으면 된다»는 판단은 대개 나중에 틀린다.
ck_run <- function(dir, tag, expr, save_raw = NULL,
                   stamp = format(Sys.Date(), "%y%m%d"), quiet = FALSE) {
  hit <- ck_find(dir, tag)
  if (!is.null(hit)) {
    if (!quiet) cat(sprintf("  [cached] %s <- %s\n", tag, basename(hit)))
    return(readRDS(hit))
  }
  out <- force(expr)
  saveRDS(out, ck_path(dir, tag, stamp))
  if (!is.null(save_raw)) saveRDS(save_raw, ck_path(dir, paste0(tag, "_raw"), stamp))
  if (!quiet) cat(sprintf("  [computed] %s\n", tag))
  out
}


# ---- 2. 작업 «안»의 replicate 체크포인트 -------------------------------------
# 셀 하나가 몇 시간이면 셀 단위 캐시만으로는 부족하다.
#
# ⚠ 반드시 «실제로 죽여서» 검증할 것. 재개 후 최종 집계가 crash 전후를 올바로 합치는지가
#   진짜 검증이고, 코드를 읽어서는 안 보인다.

#' 부분 진행 상태를 불러온다. 없으면 NULL.
part_load <- function(dir, tag) {
  p <- file.path(dir, sprintf("%s_partial.rds", tag))
  if (file.exists(p)) readRDS(p) else NULL
}

part_save <- function(dir, tag, state) {
  saveRDS(state, file.path(dir, sprintf("%s_partial.rds", tag)))
  invisible(TRUE)
}

part_clear <- function(dir, tag) {
  p <- file.path(dir, sprintf("%s_partial.rds", tag))
  if (file.exists(p)) file.remove(p)
  invisible(TRUE)
}


# ---- 3. 청크 분할 병렬 실행 --------------------------------------------------
# parLapply를 통째로 쓰지 않는다. 청크마다 클러스터를 새로 만들고 tryCatch로 감싼다.
# 워커가 죽어도 그 청크의 진행 중 항목만 잃고, 수집분은 마스터에 남는다.
#
# ⚠ 워커 수는 «이론» 아니라 «하드웨어» 상한이다. 코어가 24개여도 메모리 때문에 3개가
#   상한인 기계가 있다. 올리기 전에 실측할 것.

chunked_parlapply <- function(items, fun, ..., workers = 3L, chunk = workers,
                              on_chunk_error = NULL) {
  if (!requireNamespace("parallel", quietly = TRUE)) stop("parallel 필요")
  idx <- split(seq_along(items), ceiling(seq_along(items) / chunk))
  out <- vector("list", 0)
  for (ci in seq_along(idx)) {
    part <- items[idx[[ci]]]
    cat(sprintf("\n--- chunk %d/%d (%d items) ---\n", ci, length(idx), length(part)))
    utils::flush.console()
    got <- tryCatch({
      cl <- parallel::makeCluster(workers, type = "PSOCK")
      on.exit(try(parallel::stopCluster(cl), silent = TRUE), add = TRUE)
      r <- parallel::parLapply(cl, part, fun, ...)
      try(parallel::stopCluster(cl), silent = TRUE)
      r
    }, error = function(e) {
      cat(sprintf("  CHUNK %d FAILED: %s\n", ci, conditionMessage(e)))
      cat("  (체크포인트는 디스크에 있다. 다시 실행하면 이어받는다.)\n")
      if (is.function(on_chunk_error)) on_chunk_error(ci, e)
      NULL
    })
    if (!is.null(got)) out <- c(out, got)
    cat(sprintf("  chunk %d done; %d/%d collected\n", ci, length(out), length(items)))
    utils::flush.console()
  }
  if (length(out) < length(items))
    cat(sprintf("\n*** INCOMPLETE: %d/%d. 다시 실행하면 이어받는다. ***\n",
                length(out), length(items)))
  out
}


# ---- 4. 실행 전 점검 ---------------------------------------------------------

#' 격자의 어느 항목이 «실제로 추정에 들어가는지» 세어 보고, 안 들어가는 것을 알린다.
#' 미러링("기존 격자를 그대로 재현하면 비교 가능")은 이유가 되지 못한다.
#' @param keep  논리 벡터. TRUE인 항목만 적합/추정에 들어간다.
report_fit_relevance <- function(labels, keep, cost = NULL) {
  stopifnot(length(labels) == length(keep))
  n_drop <- sum(!keep)
  cat(sprintf("격자 %d개 중 적합에 들어가는 것 %d개, 안 들어가는 것 %d개\n",
              length(labels), sum(keep), n_drop))
  if (n_drop) {
    cat("  적합에 안 들어감:", paste(labels[!keep], collapse = ", "), "\n")
    if (!is.null(cost))
      cat(sprintf("  이들이 차지하는 비용 비중: %.0f%%\n",
                  100 * sum(cost[!keep]) / sum(cost)))
    cat("  → 그림 점으로만 쓸 것인지 확인하고, 아니면 빼라.\n")
  }
  invisible(n_drop)
}


# ---- 셀프테스트 --------------------------------------------------------------
if (!interactive() && any(grepl("--selftest", commandArgs(TRUE)))) {
  d <- file.path(tempdir(), paste0("rlr", as.integer(runif(1, 1e6, 9e6))))
  dir.create(d, showWarnings = FALSE, recursive = TRUE)
  ok <- 0L; n <- 0L
  chk <- function(lab, cond) {
    n <<- n + 1L; if (isTRUE(cond)) ok <<- ok + 1L
    cat(sprintf("  [%s] %s\n", if (isTRUE(cond)) "ok" else "FAIL", lab))
  }

  cat("== 1. 날짜 무관 캐시 ==\n")
  saveRDS(list(v = 42), ck_path(d, "cellA", stamp = "260817"))
  chk("어제 스탬프 파일을 찾는다", !is.null(ck_find(d, "cellA")))
  got <- ck_run(d, "cellA", stop("계산되면 안 된다"), quiet = TRUE)
  chk("캐시 적중 시 계산하지 않는다", identical(got$v, 42))
  chk("없는 태그는 NULL", is.null(ck_find(d, "cellZ")))

  cat("== 태그에 '.' 이 있어도 (분수 델타) ==\n")
  saveRDS(list(v = 7), ck_path(d, "S2_T1_d3.65", stamp = "260818"))
  chk("정규식 메타문자에 안 깨진다", !is.null(ck_find(d, "S2_T1_d3.65")))
  chk("접두가 다른 것을 잘못 집지 않는다", is.null(ck_find(d, "S2_T1_d3")))

  cat("== 2. 부분 체크포인트 ==\n")
  chk("없으면 NULL", is.null(part_load(d, "cellB")))
  part_save(d, "cellB", list(done = 20L))
  chk("저장/복원", identical(part_load(d, "cellB")$done, 20L))
  part_clear(d, "cellB")
  chk("정리", is.null(part_load(d, "cellB")))

  cat("== 3. 원시 draw 분리 저장 ==\n")
  ck_run(d, "cellC", list(summary = 1), save_raw = list(draws = 1:5), quiet = TRUE)
  chk("raw가 따로 저장된다", !is.null(ck_find(d, "cellC_raw")))
  chk("요약과 raw가 다른 파일", !identical(ck_find(d, "cellC"), ck_find(d, "cellC_raw")))

  cat("== 4. 적합 관련성 점검 ==\n")
  nd <- report_fit_relevance(c("a", "b", "c"), c(TRUE, FALSE, FALSE), cost = c(1, 10, 10))
  chk("빠지는 개수를 센다", identical(nd, 2L))

  cat(sprintf("\n셀프테스트 %d/%d\n", ok, n))
  unlink(d, recursive = TRUE)
  quit(status = if (ok == n) 0L else 1L)
}
