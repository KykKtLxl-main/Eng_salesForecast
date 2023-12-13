
-- 「qry11_コピーテーブル作成」「qry12_販売金額更新」
create or replace table `lixil-workspace.an1_extEng_salesForecast.t11_noukikaitoSys_motoData`
as
  with noukikaitoSys_motoData as (
    select
      cast(replace(substr(t1.hacchu_bi, 1, 10), '/', '-') as date) as hacchu_bi,
      t1.cancel_bi,
      t1.hacchu_moto,
      t1.hacchu_saki,
      t1.kanri,
      t1.num,
      t1.douji_shukka,
      t1.hacchu_naiyo_kbn,
      t1.brk_bukken_num,
      t1.brk_chumon_num,
      t1.zairyo_tehai,
      t1.zumen,
      t1.jusyo,
      t1.eigyo_kibou_syukka_bi,
      t1.dairiten_cd,
      t1.dairiten_hacchu_num,
      t1.hanbaiten_mei,
      t1.genba_mei,
      t1.nounyusaki_address,
      t1.nounyusaki_telephone,
      t1.todoke_saki,
      t1.lixil_tanto,
      t1.syohin_mei,
      t1.size,
      t1.quantity,
      t1.ryubei,
      t1.weight,
      t1.packing,
      t1.mitumori_num,

      -- 発注金額
      if(
        t1.kanri like 'F_' and t1.hacchu_naiyo_kbn in ('2','3'),
        cast(t1.eigyou_sikiri as int64) * 1.43,
        cast(t1.hanbai_kakaku as int64)
      ) as hanbai_kakaku,

      t1.syosya_tokuyaku_kakaku,
      t1.eigyo_bikou1, t1.eigyo_bikou2, t1.eigyo_bikou3, t1.eigyo_bikou4, t1.eigyo_bikou5,
      t1.distance,
      t1.fare,
      t1.nounyu_juni,
      t1.haiso_bikou,
      t1.eoc_juchu_bi,
      t1.koujou_cd,
      t1.tokuchu_cd,
      t1.nounyu_houhou,
      t1.nounyu_ire,
      t1.kigyo_noukikaitou_bi,
      t1.kigyo_noukikaitou_syukka_bi,
      t1.kigyo_noukikaitou_bikou,
      t1.eoc_noukikaitou_bi,
      t1.syukka_houhou,

      -- 確定出荷日（Nullと''が混在）
      if(
        nullif(trim(t1.kakutei_syukka_bi),'') is null ,
        '1900-01-01',   -- 日付型のカラムにNullがセット出来ないので代替
        cast(replace(substr(t1.kakutei_syukka_bi, 1, 10), '/', '-') as date)
      ) as kakutei_syukka_bi,

      t1.syukka_moto,
      t1.cancel,
      t1.syukkabi_henkou1, t1.syukkabi_henkou2, t1.syukkabi_henkou3, t1.syukkabi_henkou4, t1.syukkabi_henkou5,
      t1.chakubi_henkou1, t1.chakubi_henkou2, t1.chakubi_henkou3, t1.chakubi_henkou4, t1.chakubi_henkou5,
      t1.nounyubasyo_henkou1, t1.nounyubasyo_henkou2, t1.nounyubasyo_henkou3, t1.nounyubasyo_henkou4, t1.nounyubasyo_henkou5,
      t1.nounyuhouhou_henkou1, t1.nounyuhouhou_henkou2, t1.nounyuhouhou_henkou3, t1.nounyuhouhou_henkou4, t1.nounyuhouhou_henkou5,
      t1.kigyo_kakakukaitou_bi,
      t1.color1, t1.weight1,
      t1.color2, t1.weight2,
      t1.color3, t1.weight3,
      t1.color4, t1.weight4,
      t1.processing_cost,
      t1.yuusyo_buhin,
      t1.musyo_buhin,
      t1.syouhin_genka,
      t1.kakaku_ire,
      t1.eoc_sikiri_kaitou_bi,
      cast(t1.eigyou_sikiri as int64) as eigyou_sikiri,
      t1.cancel_hassei_hiyou,
      t1.tokki_jikou,
      t1.exc_bikou1,
      t1.exc_denpyo_syori_bi,
      t1.eoc_bikou2, t1.eoc_bikou3, t1.eoc_bikou4,
      t1.area,
      t1.zone,
      t1.horyu,
      t1.juuten_hanbai,
      t1.chaku_bi,
      t1.cyaku_jikan,
      t1.kyouyu_bikou1, t1.kyouyu_bikou2,

      t2.OFFICE_CD, t2.OFFICE_NAME,
      t2.N_LEVEL_CD, t2.N_LEVEL_NAME,
      t2.F_LEVEL_CD, t2.F_LEVEL_NAME,
      t2.B_LEVEL_CD, t2.B_LEVEL_NAME,
      t2.J_LEVEL_CD, t2.J_LEVEL_NAME
    from `lixil-workspace.an1_extEng_salesForecast.t01_noukikaitouSystem_data` as t1

    -- 「発注元」が正しく入力されていないパターンが見つかったのでleft joinにする
    left join `lixil-dwh.pii_an1_bo.M_NEW_ORG_SYSTEM` as t2 on t1.hacchu_moto = t2.OFFICE_CD

    where t1.cancel_bi is null
  )

  select *
  from noukikaitoSys_motoData
  order by hacchu_bi
;


