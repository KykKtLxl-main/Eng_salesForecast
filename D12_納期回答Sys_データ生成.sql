
declare pre_month_lastday, cur_month_lastday, nxt1_month_lastday, nxt2_month_lastday date;
declare today date;
set pre_month_lastday = (
  date(
    extract(year from current_date('Asia/Tokyo')),
    extract(month from current_date('Asia/Tokyo')),
    1
  ) -1
);
set cur_month_lastday = (
  date(
    extract(year from date_add(current_date('Asia/Tokyo'), interval 1 month)),
    extract(month from date_add(current_date('Asia/Tokyo'), interval 1 month)),
    1
  ) -1
);
set nxt1_month_lastday = (
  date(
    extract(year from date_add(current_date('Asia/Tokyo'), interval 2 month)),
    extract(month from date_add(current_date('Asia/Tokyo'), interval 2 month)),
    1
  ) -1
);
set nxt2_month_lastday = (
  date(
    extract(year from date_add(current_date('Asia/Tokyo'), interval 3 month)),
    extract(month from date_add(current_date('Asia/Tokyo'), interval 3 month)),
    1
  ) -1
);
set today = current_date('Asia/Tokyo');

-- 当月全FAX注文
create or replace table `lixil-workspace.an1_extEng_salesForecast.t12_noukikaitouSys_curMonth_FAXorder` as (
  select
    *
  from `lixil-workspace.an1_extEng_salesForecast.t11_noukikaitouSys_motoData`
  where kakutei_syukka_bi between pre_month_lastday and cur_month_lastday
)
;
-- 当月注残
create or replace table `lixil-workspace.an1_extEng_salesForecast.t12_noukikaitouSys_curMonth_backlog` as (
  select *
  from `lixil-workspace.an1_extEng_salesForecast.t12_noukikaitouSys_curMonth_FAXorder`
  where kakutei_syukka_bi < today and eoc_denpyo_syori_bi != today
)
;

-- 翌月納期確定