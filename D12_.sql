-- 2023-12-13 16:59:25
-- 当月全FAX注文
declare pre_month_lastday, cur_month_lastday date;
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

select
*
from `lixil-workspace.an1_extEng_salesForecast.t11_noukikaitoSys_motoData`
where kakutei_syukka_bi between pre_month_lastday and cur_month_lastday
;