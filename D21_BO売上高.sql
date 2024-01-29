
create or replace table `lixil-workspace.an1_extEng_salesForecast.t10_noukikaitouSystem_data` as (
  select * from `lixil-workspace.an1_extEng_salesForecast.t01_noukikaitouSystem_data`
)


create or replace table `lixil-workspace.an1_extEng_salesForecast.t21_salesAmount_fromBO` as (

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
      sum(t1.QUANTITY) as quantity
    from `lixil-dwh.pii_an1_bo.T_CSP_BASE_DATA` as t1   --BO CSP基礎データ(予実)
    join `lixil-dwh.pii_an1_bo.M_NEW_ORG_SYSTEM` as t2 on t1.OFFICE_CD = t2.OFFICE_CD
    join `lixil-dwh.pii_an1_bo.M_GOODS_CLASS_SYSTEM` as t3 on t1.GOODS_CLASS_TOTAL = t3.GOODS_CLASS_TOTAL
    where t1.KIKAN_KBN = 'T'        --T：TRAIN /S：SIS
      and t1.LCR_CD <> '1'          --LCR区分は「1」を除かないと集計結果が重複する
      and t1.DOUBLE_DEALING <> '1'  --二重売（代納店への部材支給）を除外する
      and t2.V_LEVEL_CD = 'V00100'
      and t3.T_GOODS_CLASS_CD = 'T41334'

      -- 当期を含む３年度前（当期、前期、前々期）を出力対象とする
      and (
        t1.YEAR_MONTH >= if (
          extract(month from CURRENT_DATE('Asia/Tokyo')) between 4 and 12,
          extract(year from CURRENT_DATE('Asia/Tokyo')) -2,
          extract(year from CURRENT_DATE('Asia/Tokyo')) -3
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

  select * from salesAmount_fromBO
)
;
