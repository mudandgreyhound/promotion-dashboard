-- v_bill_summary / v_promotion_summary (แก้ 2026-11-23)
-- คีย์บิล = (ShopCode, BillNo); BILLPAYTB tendtype=MEM ยุบเหลือ 1 แถว/บิลด้วย LATERAL (เดิม join ด้วย BillNo อย่างเดียว แถวงอก 2-3 เท่า)
-- MemCode = BillHdrTB.Remark (ตรง MemberTB 98%) สำรองด้วย CardNo ที่ไม่ว่าง; MemberTypeName สำรองด้วย tier จาก MemberTB
-- v_promotion_summary เพิ่มคอลัมน์ ShopCode ท้ายสุด (คอลัมน์เดิมชื่อ/ลำดับเหมือนเดิม)
CREATE OR REPLACE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `v_bill_summary` AS
SELECT h.ShopCode AS ShopCode, s.ShopDes1 AS ShopName, h.BillDate AS BillDate, h.BillNo AS BillNo,
       h.BillTime AS BillTime, h.TableNo AS TableNo, h.NoPerson AS NoPerson, h.SubTotal AS SubTotal,
       h.MemDisc AS MemDisc, h.BillDisc AS BillDisc, h.OthChrg AS ServiceCharge, h.NetBAmt AS NetBAmt,
       h.VoidBill AS VoidBill,
       CASE WHEN p.n > 0 THEN COALESCE(NULLIF(TRIM(h.Remark), ''), p.CardNo) END AS MemCode,
       p.CardName AS MemberName, p.Prefix AS MemType,
       COALESCE(mt.MemTNam1, mt2.MemTNam1) AS MemberTypeName,
       CASE WHEN p.n > 0 THEN 'Member' ELSE 'Guest' END AS CustomerType
FROM BillHdrTB h
LEFT JOIN PosShopTB s ON s.ShopCode = h.ShopCode
LEFT JOIN LATERAL (
    SELECT COUNT(*) AS n, MAX(NULLIF(x.CardNo, '')) AS CardNo,
           MAX(CASE WHEN x.CardNo <> '' THEN x.CardName END) AS CardName, MAX(x.Prefix) AS Prefix
    FROM BILLPAYTB x
    WHERE x.ShopCode = h.ShopCode AND x.BillNo = h.BillNo AND x.tendtype = 'MEM') p ON TRUE
LEFT JOIN MemTypetb mt ON mt.MemType = p.Prefix
LEFT JOIN MemberTB m ON p.n > 0 AND m.MemCode = COALESCE(NULLIF(TRIM(h.Remark), ''), p.CardNo)
LEFT JOIN MemTypetb mt2 ON mt2.MemType = m.MemType
WHERE (h.VoidBill IS NULL OR h.VoidBill = 0);

CREATE OR REPLACE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `v_promotion_summary` AS
SELECT disc.BillDate AS BillDate, disc.BillNo AS BillNo, disc.DiscCode AS DiscCode, disc.DiscName AS DiscName,
       disc.DiscType AS DiscType, disc.DiscLevel AS DiscLevel, disc.ItemCode AS ItemCode, i.ItemDes1 AS ItemDes1,
       disc.DiscAmt AS DiscAmt, disc.TotBillDisc AS TotBillDisc,
       COALESCE(mt.MemTNam1, mt2.MemTNam1) AS MemberTypeName,
       disc.ShopCode AS ShopCode
FROM ONLBillMultiDiscTB disc
LEFT JOIN ItmMasTB i ON i.ItemCode = disc.ItemCode
LEFT JOIN LATERAL (
    SELECT MAX(x.Prefix) AS Prefix FROM BILLPAYTB x
    WHERE x.ShopCode = disc.ShopCode AND x.BillNo = disc.BillNo AND x.tendtype = 'MEM') p ON TRUE
LEFT JOIN MemTypetb mt ON mt.MemType = p.Prefix
LEFT JOIN LATERAL (
    SELECT MAX(NULLIF(TRIM(hh.Remark), '')) AS MemCode FROM BillHdrTB hh
    WHERE hh.ShopCode = disc.ShopCode AND hh.BillNo = disc.BillNo) hb ON TRUE
LEFT JOIN MemberTB m ON m.MemCode = hb.MemCode
LEFT JOIN MemTypetb mt2 ON mt2.MemType = m.MemType;
