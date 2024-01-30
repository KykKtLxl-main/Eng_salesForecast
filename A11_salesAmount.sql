
-- STEP1：BO「T_CSP_BASE_DATA（予実）」から実績データを出力
create or replace table `lixil-workspace.an1_extEng_salesForecast.A11_salesAmount_fromBO` as (
  with salesAmount_fromBO as (
    select
      t1.YEAR_MONTH,
      t1.KIKAN_KBN,
      t1.BDB_CD,
      t1.LCR_CD, t1.LCR_KBN,
      t1.DOUBLE_DEALING,
      t2.V_LEVEL_CD, t2.V_LEVEL_NAME,
      t2.P_LEVEL_CD, t2.P_LEVEL_NAME,
      t2.N_LEVEL_CD, t2.N_LEVEL_NAME,
      t2.F_LEVEL_CD, t2.F_LEVEL_NAME,
      t2.B_LEVEL_CD, t2.B_LEVEL_NAME,
      t2.J_LEVEL_CD, t2.J_LEVEL_NAME,
      t1.OFFICE_CD, t2.OFFICE_NAME,

      t3.L_GOODS_CLASS_CD, t3.L_GOODS_CLASS_NAME,
      t3.T_GOODS_CLASS_CD, t3.T_GOODS_CLASS_NAME,
      t3.S_GOODS_CLASS_CD, t3.S_GOODS_CLASS_NAME,
      t3.Q_GOODS_CLASS_CD, t3.Q_GOODS_CLASS_NAME,
      t3.P_GOODS_CLASS_CD, t3.P_GOODS_CLASS_NAME,
      t3.N_GOODS_CLASS_CD, t3.N_GOODS_CLASS_NAME,
      t3.M_GOODS_CLASS_CD, t3.M_GOODS_CLASS_NAME,

      sum(t1.AMOUNT) as amount,
      -- sum(t1.QUANTITY) as quantity
    from `lixil-dwh.pii_an1_bo.T_CSP_BASE_DATA` as t1   --BO CSP基礎データ(予実)
    join `lixil-dwh.pii_an1_bo.M_NEW_ORG_SYSTEM` as t2 on t1.OFFICE_CD = t2.OFFICE_CD
    join `lixil-dwh.pii_an1_bo.M_GOODS_CLASS_SYSTEM` as t3 on t1.GOODS_CLASS_TOTAL = t3.GOODS_CLASS_TOTAL
    where t1.KIKAN_KBN = 'T'        --T：TRAIN /S：SIS
      and t1.LCR_CD <> '1'          --LCR区分は「1」を除かないと集計結果が重複する
      and t1.DOUBLE_DEALING <> '1'  --二重売（代納店への部材支給）を除外する
      and t2.V_LEVEL_CD = 'V00100'
      and t3.T_GOODS_CLASS_CD = 'T41334'

      -- 当期を含む４年度前（当期、前期、前々期、前々々期）を出力対象とする（前々々期は前年比較用）
      and (
        t1.YEAR_MONTH >=
        if (
          extract(month from CURRENT_DATE('Asia/Tokyo')) between 4 and 12,
          extract(year from CURRENT_DATE('Asia/Tokyo')) -3,
          extract(year from CURRENT_DATE('Asia/Tokyo')) -4
        ) * 100 + 4
        and
        t1.YEAR_MONTH <= extract(year from CURRENT_DATE('Asia/Tokyo')) *100 + extract(month from CURRENT_DATE('Asia/Tokyo'))
      )

    group by
      t1.YEAR_MONTH,
      t1.KIKAN_KBN,
      t1.BDB_CD,
      t1.LCR_CD, t1.LCR_KBN,
      t1.DOUBLE_DEALING,
      t2.V_LEVEL_CD, t2.V_LEVEL_NAME,
      t2.P_LEVEL_CD, t2.P_LEVEL_NAME,
      t2.N_LEVEL_CD, t2.N_LEVEL_NAME,
      t2.F_LEVEL_CD, t2.F_LEVEL_NAME,
      t2.B_LEVEL_CD, t2.B_LEVEL_NAME,
      t2.J_LEVEL_CD, t2.J_LEVEL_NAME,
      t1.OFFICE_CD, t2.OFFICE_NAME,

      t3.L_GOODS_CLASS_CD, t3.L_GOODS_CLASS_NAME,
      t3.T_GOODS_CLASS_CD, t3.T_GOODS_CLASS_NAME,
      t3.S_GOODS_CLASS_CD, t3.S_GOODS_CLASS_NAME,
      t3.Q_GOODS_CLASS_CD, t3.Q_GOODS_CLASS_NAME,
      t3.P_GOODS_CLASS_CD, t3.P_GOODS_CLASS_NAME,
      t3.N_GOODS_CLASS_CD, t3.N_GOODS_CLASS_NAME,
      t3.M_GOODS_CLASS_CD, t3.M_GOODS_CLASS_NAME
  )

  select

    -- FYE追加
    if(
      cast(right(cast(YEAR_MONTH as string),2) as int64) < 4 ,
      cast(left(cast(YEAR_MONTH as string),4) as int64),
      cast(left(cast(YEAR_MONTH as string),4) as int64) +1,
    ) as FYE,

    *
  from salesAmount_fromBO
)
;

-- STEP2：前年値を横に並べるテーブルを作成
-- （１）当年が0円だと表示されない ⇒cross joinで全組み合わせを軸にしたテーブルを作成
create or replace table `lixil-workspace.an1_extEng_salesForecast.A12_salesAmount_addPre` as (

  with YearMonth_range as (
      select year_month
      from unnest(
          generate_date_array(
              date(
                  if( extract(month from CURRENT_DATE('Asia/Tokyo')) between 4 and 12,
                      extract(year from CURRENT_DATE('Asia/Tokyo')) -2,
                      extract(year from CURRENT_DATE('Asia/Tokyo')) -3),
              4,1),
              CURRENT_DATE('Asia/Tokyo')
          )
      ) as year_month
      where extract(day FROM year_month) = 1
      order by year_month
  ),
  yearMonth_list as (
      select extract(year from year_month) *100 + extract(month from year_month) as YEAR_MONTH
      from YearMonth_range
  ),

  salesOffice as (
    select distinct
      V_LEVEL_CD, V_LEVEL_NAME,
      P_LEVEL_CD, P_LEVEL_NAME,
      N_LEVEL_CD, N_LEVEL_NAME,
      F_LEVEL_CD, F_LEVEL_NAME,
      B_LEVEL_CD, B_LEVEL_NAME,
      D_LEVEL_CD, D_LEVEL_NAME,
      J_LEVEL_CD, J_LEVEL_NAME,
      OFFICE_CD, OFFICE_NAME
    from `lixil-dwh.pii_an1_bo.M_NEW_ORG_SYSTEM`
    where YEAR_MONTH = cast(format_date('%Y%m',CURRENT_DATE('Asia/Tokyo')) as int64)
    and V_LEVEL_CD = 'V00100'
  ),

  goodsClass_byM as (
    select distinct
      T_GOODS_CLASS_CD, T_GOODS_CLASS_NAME,
      S_GOODS_CLASS_CD, S_GOODS_CLASS_NAME,
      Q_GOODS_CLASS_CD, Q_GOODS_CLASS_NAME,
      P_GOODS_CLASS_CD, P_GOODS_CLASS_NAME,
      N_GOODS_CLASS_CD, N_GOODS_CLASS_NAME,
      M_GOODS_CLASS_CD, M_GOODS_CLASS_NAME
    from `lixil-dwh.pii_an1_bo.M_GOODS_CLASS_SYSTEM`
    where T_GOODS_CLASS_CD = 'A11334'
  )

  select
    *,
    cast(0 as numeric) as amount,
    cast(0 as numeric) as pre_amount
  from yearMonth_list
  cross join salesOffice
  cross join goodsClass_byM
);

--（２）実績データを一時テーブルにする
create or replace temp table salesAmount_extEng as (
  select YEAR_MONTH, OFFICE_CD, M_GOODS_CLASS_CD, sum(amount) as amount
  from `lixil-workspace.an1_extEng_salesForecast.A11_salesAmount_fromBO`
  group by YEAR_MONTH, OFFICE_CD, M_GOODS_CLASS_CD
);
--（３）当年実績を追加
update `lixil-workspace.an1_extEng_salesForecast.A12_salesAmount_addPre` as t1
set t1.amount = t2.amount
from salesAmount_extEng as t2
where t1.YEAR_MONTH = t2.YEAR_MONTH
  and t1.OFFICE_CD = t2.OFFICE_CD
  and t1.M_GOODS_CLASS_CD = t2.M_GOODS_CLASS_CD
;
--（４）前年実績を追加
update `lixil-workspace.an1_extEng_salesForecast.A12_salesAmount_addPre` as t1
set t1.pre_amount = t2.amount
from salesAmount_extEng as t2
where t1.YEAR_MONTH = t2.YEAR_MONTH +100
  and t1.OFFICE_CD = t2.OFFICE_CD
  and t1.M_GOODS_CLASS_CD = t2.M_GOODS_CLASS_CD
;
--（５）当年値+前年値=0円ならばデータ削除
delete from `lixil-workspace.an1_extEng_salesForecast.A12_salesAmount_addPre`
where amount + pre_amount = 0
;

--
