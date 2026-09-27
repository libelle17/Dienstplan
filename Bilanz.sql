-- Jahresbilanzen (Urlaub, Urlaubsstunden, Überstunden, Fortbildung, Planstunden) in einem SQL-Befehl,
-- bildet gesBilanz/EinzelBilanz/UrlAnspr (VB) nach. Kumulierte Werte je 31.12. bzw. Austrittstag.
-- Urlaub in Tagen = Urlaubsstunden / (0,2 * WAZ) (vor 2020 und in der Ausbildung WAZ = 38,5)
WITH RECURSIVE
ma AS (SELECT m.persnr,
        COALESCE((SELECT MIN(ab) FROM wochenplan w WHERE w.persnr=m.persnr),DATE(20040701)) von,
        IF(m.aus>0,m.aus,MAKEDATE(YEAR(CURDATE())+4,1)-INTERVAL 1 DAY) bis,
        IF(m.ausbende>0,m.ausbende,DATE(19000101)) ausbende
       FROM mitarbeiter m),
wp AS (SELECT w.*,COALESCE(LEAD(ab) OVER (PARTITION BY persnr ORDER BY ab),DATE(99991231)) nab FROM wochenplan w WHERE ab<>0),
-- Urlaubsanspruch in Stunden je Jahr (wie UrlAnspr, iru=0)
jr AS (SELECT persnr,YEAR(von) jahr,von,bis,ausbende FROM ma UNION ALL SELECT persnr,jahr+1,von,bis,ausbende FROM jr WHERE jahr<YEAR(bis)),
jr2 AS (SELECT jr.*,MAKEDATE(jahr,1) gv,MAKEDATE(jahr+1,1) gb,
         (GREATEST(von,MAKEDATE(jahr,1))<DATE(20200101) OR GREATEST(von,MAKEDATE(jahr,1))<ausbende) ntg FROM jr),
seg AS (SELECT w.persnr,w.ab,w.waz,w.urlaub,
         COALESCE(w.nab2,(SELECT IF(aus>0,aus,DATE(99991231)) FROM mitarbeiter WHERE persnr=w.persnr)) bis,
         COALESCE(IF(w.waz=0,(SELECT waz FROM wochenplan x WHERE x.persnr=w.persnr AND x.waz<>0 AND x.ab<w.ab ORDER BY x.ab DESC LIMIT 1),w.waz),38.5) rwaz
        FROM (SELECT wochenplan.*,LEAD(ab) OVER (PARTITION BY persnr ORDER BY ab) nab2 FROM wochenplan) w),
ansp AS (SELECT j.persnr,j.jahr,
          SUM(DATEDIFF(LEAST(s.bis,j.gb),GREATEST(s.ab,j.gv))/DATEDIFF(j.gb,j.gv)*IF(j.ntg,38.5,s.rwaz)*0.2*s.urlaub) uaah
         FROM jr2 j JOIN seg s ON s.persnr=j.persnr AND s.bis>j.gv AND s.ab<j.gb
         GROUP BY j.persnr,j.jahr),
-- Tage
tg AS (SELECT ma.persnr,ma.bis,ma.ausbende,ma.von + INTERVAL s.seq DAY tag FROM ma JOIN seq_0_to_40000 s ON s.seq<=DATEDIFF(ma.bis,ma.von)),
t1 AS (SELECT tg.persnr,tg.tag,tg.bis,
        (tg.tag<DATE(20200101) OR tg.tag<tg.ausbende) ntg,
        IF(wp.waz=0,38.5,wp.waz) waz,
        IF(obft(tg.tag)>0 OR DATE_FORMAT(tg.tag,'%m%d')='1231','WF',  -- Feiertag (VB inkl. Silvester)
           CASE DAYOFWEEK(tg.tag) WHEN 1 THEN wp.so WHEN 2 THEN wp.mo WHEN 3 THEN wp.di WHEN 4 THEN wp.mi
                                  WHEN 5 THEN wp.do WHEN 6 THEN wp.fr ELSE wp.sa END) vgb,
        d.artnr dp,
        COALESCE(z.ausbez,0) ausbez, COALESCE(z.urlhaus,0) urlhaus,
        (tg.tag=MAKEDATE(YEAR(tg.tag),1) OR tg.tag=(SELECT von FROM ma WHERE ma.persnr=tg.persnr)) jahranfang
       FROM tg
       LEFT JOIN wp ON wp.persnr=tg.persnr AND tg.tag>=wp.ab AND tg.tag<wp.nab
       LEFT JOIN dienstplan d ON d.persnr=tg.persnr AND d.tag=tg.tag
       LEFT JOIN (SELECT persnr,tag,SUM(ausbez) ausbez,SUM(urlhaus) urlhaus FROM ausbez GROUP BY persnr,tag) z ON z.persnr=tg.persnr AND z.tag=tg.tag),
t2 AS (SELECT t1.*,
        IF(TRIM(vgb) REGEXP '^[0-9]+(,[0-9]+)?$',REPLACE(TRIM(vgb),',','.')+0,0) vstd,
        IF(TRIM(dp) REGEXP '^[0-9]+(,[0-9]+)?$',REPLACE(TRIM(dp),',','.')+0,NULL) dstd,
        BINARY dp IN ('g','u','uw') AND NOT (ntg AND BINARY vgb='-' AND BINARY dp='g') urltag
       FROM t1),
t3 AS (SELECT t2.persnr,t2.tag,t2.bis,t2.ntg,t2.waz,t2.vstd,
        -vstd + CASE WHEN dp IS NULL OR LENGTH(dp)=0 THEN vstd
                     WHEN BINARY dp IN ('b','hFT','WF','k','ki','f','fw','g','u','uw','su') THEN vstd
                     ELSE COALESCE(dstd,0) END - ausbez ue,
        (BINARY dp IN ('f','fw')) fb,
        IF(urltag,IF(ntg,7.7,vstd),0) + urlhaus - IF(jahranfang,COALESCE(a.uaah,0),0) uh
       FROM t2 LEFT JOIN ansp a ON a.persnr=t2.persnr AND a.jahr=YEAR(t2.tag))
SELECT persnr,YEAR(tag) jahr,
       ROUND(uh_kum/(0.2*IF(ntg,38.5,waz)),1) urlaub, ROUND(uh_kum,1) urlstd,
       ROUND(ue_kum,1) ueberstunden, ROUND(fb_kum,1) fortbildung, ROUND(pst_kum,1) planstunden
FROM (SELECT persnr,tag,bis,ntg,waz,
       SUM(uh)   OVER w uh_kum,
       SUM(ue)   OVER w ue_kum,
       SUM(fb)   OVER w fb_kum,
       SUM(vstd) OVER w pst_kum
      FROM t3 WINDOW w AS (PARTITION BY persnr ORDER BY tag)) k
WHERE DATE_FORMAT(tag,'%m%d')='1231' OR tag=bis
ORDER BY persnr,jahr;
