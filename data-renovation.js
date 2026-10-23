// 添付Excel「見積りベース3 - コピー.xlsx」から生成した改修用価格データ。
// 現行価格は各シート左端の価格ブロック。0・空欄は未登録として null のまま保持する。
const RENOVATION_CATEGORIES = [
  {
    "id": "cat-0",
    "label": "DS",
    "sourceLabel": "30%up R8.7.1(確定価格)"
  },
  {
    "id": "cat-1-vt",
    "label": "高圧VT",
    "sourceLabel": "定価"
  },
  {
    "id": "cat-1-ct-relay",
    "label": "高圧CT(継電器用)",
    "sourceLabel": "定価"
  },
  {
    "id": "cat-1-ct-meter",
    "label": "高圧CT(指示計器用)",
    "sourceLabel": "定価"
  },
  {
    "id": "cat-2-vcb",
    "label": "VCB",
    "sourceLabel": "30%up R8.7.1(確定価格)"
  },
  {
    "id": "cat-2-contactor",
    "label": "コンタクタ",
    "sourceLabel": "30%up R8.7.1(確定価格)"
  },
  {
    "id": "cat-3",
    "label": "継電器",
    "sourceLabel": "三菱CTD 30%up・光商工 R8.7.1"
  },
  {
    "id": "cat-4",
    "label": "太陽光",
    "sourceLabel": "各々UP R6.10.1"
  },
  {
    "id": "cat-5",
    "label": "ﾏﾙﾁ",
    "sourceLabel": "45%up R6.10.1"
  },
  {
    "id": "cat-6",
    "label": "LBS",
    "sourceLabel": "30%up R8.7.1(確定価格)"
  },
  {
    "id": "cat-7",
    "label": "PC・LA",
    "sourceLabel": "定価"
  },
  {
    "id": "cat-8",
    "label": "SC・SR",
    "sourceLabel": "SR15％ SC７％ 2026.7"
  },
  {
    "id": "cat-9",
    "label": "TR",
    "sourceLabel": "日立"
  },
  {
    "id": "cat-10",
    "label": "ELR",
    "sourceLabel": "定価"
  },
  {
    "id": "cat-11",
    "label": "低圧VT・CT",
    "sourceLabel": "定価"
  },
  {
    "id": "cat-12",
    "label": "ﾒｰﾀｰ",
    "sourceLabel": "定価"
  },
  {
    "id": "cat-13",
    "label": "MCB(表)",
    "sourceLabel": "1.3"
  },
  {
    "id": "cat-14",
    "label": "MCB(裏) ",
    "sourceLabel": "30%up R8.7.1"
  },
  {
    "id": "cat-15",
    "label": "MCB(FP)",
    "sourceLabel": "30%up R8.7.1"
  },
  {
    "id": "cat-16",
    "label": "ELB(表)",
    "sourceLabel": "30%up R8.7.1"
  },
  {
    "id": "cat-17",
    "label": "ELB(裏)",
    "sourceLabel": "30%up R8.7.1"
  },
  {
    "id": "cat-18",
    "label": "ELB(FP)",
    "sourceLabel": "30%up R8.7.1"
  },
  {
    "id": "cat-19",
    "label": "ｽﾘﾑ・KC",
    "sourceLabel": "30%up R8.7.1"
  },
  {
    "id": "cat-20",
    "label": "特殊Br・DCS",
    "sourceLabel": "15%up"
  },
  {
    "id": "cat-21",
    "label": "Panasonic",
    "sourceLabel": "黄色の箇所 2026/2/6 定価・仕入れとも約1.1倍"
  },
  {
    "id": "cat-misc",
    "label": "その他",
    "sourceLabel": "個別記録"
  }
];

const RENOVATION_ITEMS = [
  {
    "id": "REN-01-004",
    "categoryId": "cat-0",
    "category": "DS",
    "name": "DS 7.2kV 200A 単極単投",
    "spec": "",
    "defaultQty": 3,
    "listPrice": 21400,
    "netPrice": 11770,
    "costPrice": 7970,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 16450,
        "netPrice": 9100,
        "costPrice": 6110
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 14300,
        "netPrice": 8000,
        "costPrice": 5040
      }
    ]
  },
  {
    "id": "REN-01-005",
    "categoryId": "cat-0",
    "category": "DS",
    "name": "DS 7.2kV 400A 単極単投",
    "spec": "",
    "defaultQty": 3,
    "listPrice": 24800,
    "netPrice": 13640,
    "costPrice": 9150,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 19090,
        "netPrice": 10500,
        "costPrice": 7060
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 16600,
        "netPrice": 9000,
        "costPrice": 6170
      }
    ]
  },
  {
    "id": "REN-01-006",
    "categoryId": "cat-0",
    "category": "DS",
    "name": "DS 7.2kV 600A 単極単投",
    "spec": "",
    "defaultQty": 3,
    "listPrice": 37700,
    "netPrice": 20735,
    "costPrice": 13870,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 29000,
        "netPrice": 16000,
        "costPrice": 10660
      }
    ]
  },
  {
    "id": "REN-01-007",
    "categoryId": "cat-0",
    "category": "DS",
    "name": "DS 7.2kV 200A 三極単投",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 125500,
    "netPrice": 69025,
    "costPrice": 46260,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 96500,
        "netPrice": 53100,
        "costPrice": 35580
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 83900,
        "netPrice": 46000,
        "costPrice": 31440
      }
    ]
  },
  {
    "id": "REN-01-008",
    "categoryId": "cat-0",
    "category": "DS",
    "name": "DS 7.2kV 400A 三極単投",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 150500,
    "netPrice": 82775,
    "costPrice": 55640,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 115800,
        "netPrice": 63700,
        "costPrice": 42780
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 100700,
        "netPrice": 55000,
        "costPrice": 37800
      }
    ]
  },
  {
    "id": "REN-02-003",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "VT 6600/110V 50VA",
    "spec": "PD-50HF",
    "defaultQty": 2,
    "listPrice": 62900,
    "netPrice": 34600,
    "costPrice": 25200,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-004",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "VT 6600/110V 100VA",
    "spec": "PD-100HF",
    "defaultQty": 2,
    "listPrice": 77100,
    "netPrice": 42500,
    "costPrice": 30800,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-005",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "VT 6600/110V 200VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 163800,
    "netPrice": 91000,
    "costPrice": 52500,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-006",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "VT 380/110V 50VA",
    "spec": "PD-50HF",
    "defaultQty": 2,
    "listPrice": 84000,
    "netPrice": 46200,
    "costPrice": 31200,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-007",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "VT 380/110V 15VA",
    "spec": "PE-15F",
    "defaultQty": 2,
    "listPrice": 56000,
    "netPrice": 30800,
    "costPrice": 20740,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-009",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "OPTr 6600/110V 300VA",
    "spec": "EMT-K",
    "defaultQty": 1,
    "listPrice": 145000,
    "netPrice": 80000,
    "costPrice": 53890,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-010",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "OPTr 6600/110V 600VA",
    "spec": "EMT-BB",
    "defaultQty": 1,
    "listPrice": 314000,
    "netPrice": 173000,
    "costPrice": 107400,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-013",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 10/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 94000,
    "netPrice": 52000,
    "costPrice": 34970,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-014",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 15/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 94000,
    "netPrice": 52000,
    "costPrice": 34970,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-015",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 20/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 94000,
    "netPrice": 52000,
    "costPrice": 34970,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-016",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 25/5A 25VA",
    "spec": "受注品",
    "defaultQty": 2,
    "listPrice": 94000,
    "netPrice": 52000,
    "costPrice": 34970,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-017",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 30/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 94000,
    "netPrice": 52000,
    "costPrice": 34970,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-018",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 40/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 94000,
    "netPrice": 52000,
    "costPrice": 34970,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-019",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 50/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 94000,
    "netPrice": 52000,
    "costPrice": 34970,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-020",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 60/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 94000,
    "netPrice": 52000,
    "costPrice": 34970,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-021",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 75/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 94000,
    "netPrice": 52000,
    "costPrice": 34970,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-022",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 80/5A 25VA",
    "spec": "受注品",
    "defaultQty": 2,
    "listPrice": 94000,
    "netPrice": 52000,
    "costPrice": 34970,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-023",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 100/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 94000,
    "netPrice": 52000,
    "costPrice": 34970,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-024",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 120/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 94000,
    "netPrice": 52000,
    "costPrice": 34970,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-025",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 150/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 109000,
    "netPrice": 60000,
    "costPrice": 40630,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-026",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 200/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 109000,
    "netPrice": 60000,
    "costPrice": 40630,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-027",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 250/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 126000,
    "netPrice": 70000,
    "costPrice": 46670,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-028",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 300/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 126000,
    "netPrice": 70000,
    "costPrice": 46670,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-029",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 400/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 126000,
    "netPrice": 70000,
    "costPrice": 46670,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-030",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 500/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 126000,
    "netPrice": 70000,
    "costPrice": 46670,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-034",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 5/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 31400,
    "netPrice": 18000,
    "costPrice": 11640,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-035",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 10/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 31400,
    "netPrice": 18000,
    "costPrice": 11640,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-036",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 15/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 31400,
    "netPrice": 18000,
    "costPrice": 11640,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-037",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 20/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 31400,
    "netPrice": 18000,
    "costPrice": 11640,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-038",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 25/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 31400,
    "netPrice": 18000,
    "costPrice": 11640,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-039",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 30/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 31400,
    "netPrice": 18000,
    "costPrice": 11640,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-040",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 40/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 31400,
    "netPrice": 18000,
    "costPrice": 11640,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-041",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 50/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 31400,
    "netPrice": 18000,
    "costPrice": 11640,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-042",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 60/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 31400,
    "netPrice": 18000,
    "costPrice": 11640,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-043",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 75/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 31400,
    "netPrice": 18000,
    "costPrice": 11640,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-044",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 80/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 31400,
    "netPrice": 18000,
    "costPrice": 11640,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-045",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 100/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 31400,
    "netPrice": 18000,
    "costPrice": 11640,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-046",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 120/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 32800,
    "netPrice": 19000,
    "costPrice": 12160,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-047",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 150/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 32800,
    "netPrice": 19000,
    "costPrice": 12160,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-048",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 200/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 33000,
    "netPrice": 19000,
    "costPrice": 12220,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-049",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 250/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 34300,
    "netPrice": 19000,
    "costPrice": 12740,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-050",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 300/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 34300,
    "netPrice": 19000,
    "costPrice": 12740,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-051",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 400/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 34500,
    "netPrice": 19000,
    "costPrice": 12810,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-052",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 500/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 36100,
    "netPrice": 20000,
    "costPrice": 13390,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-053",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 600/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 39300,
    "netPrice": 22000,
    "costPrice": 14560,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-02-054",
    "categoryId": "cat-1",
    "category": "高圧VT・CT",
    "name": "CT 6.9kV 700/5A 25VA",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 39300,
    "netPrice": 22000,
    "costPrice": 14560,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-03-004",
    "categoryId": "cat-2",
    "category": "VCB・ｺﾝﾀｸﾀ",
    "name": "VCB 7.2kV 400A 手動",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 449000,
    "netPrice": 233220,
    "costPrice": 117315,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 345000,
        "netPrice": 179400,
        "costPrice": 90060
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 300000,
        "netPrice": 156000,
        "costPrice": 78210
      }
    ]
  },
  {
    "id": "REN-03-005",
    "categoryId": "cat-2",
    "category": "VCB・ｺﾝﾀｸﾀ",
    "name": "VCB 7.2kV 400A 手動 電圧引外し",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 504000,
    "netPrice": 261690,
    "costPrice": 149040,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 387000,
        "netPrice": 201300,
        "costPrice": 114610
      }
    ]
  },
  {
    "id": "REN-03-006",
    "categoryId": "cat-2",
    "category": "VCB・ｺﾝﾀｸﾀ",
    "name": "VCB 7.2kV 400A 電動",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 647000,
    "netPrice": 336440,
    "costPrice": 207000,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 497000,
        "netPrice": 258800,
        "costPrice": 159000
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 432000,
        "netPrice": 225000,
        "costPrice": 165600
      }
    ]
  },
  {
    "id": "REN-03-007",
    "categoryId": "cat-2",
    "category": "VCB・ｺﾝﾀｸﾀ",
    "name": "VCB 7.2kV 600A 手動",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 665000,
    "netPrice": 345410,
    "costPrice": 170890,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 511000,
        "netPrice": 265700,
        "costPrice": 136500
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 444000,
        "netPrice": 231000,
        "costPrice": 113930
      }
    ]
  },
  {
    "id": "REN-03-008",
    "categoryId": "cat-2",
    "category": "VCB・ｺﾝﾀｸﾀ",
    "name": "VCB 7.2kV 600A 手動 電圧引外し",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 718000,
    "netPrice": 373100,
    "costPrice": 206570,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 552000,
        "netPrice": 287000,
        "costPrice": 158900
      }
    ]
  },
  {
    "id": "REN-03-009",
    "categoryId": "cat-2",
    "category": "VCB・ｺﾝﾀｸﾀ",
    "name": "VCB 7.2kV 600A 電動",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 853000,
    "netPrice": 444080,
    "costPrice": 269500,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 656000,
        "netPrice": 341600,
        "costPrice": 207000
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 570000,
        "netPrice": 297000,
        "costPrice": 216000
      }
    ]
  },
  {
    "id": "REN-03-016",
    "categoryId": "cat-2",
    "category": "VCB・ｺﾝﾀｸﾀ",
    "name": "VCS 6.6kV 200A 常時励磁 ACDC100V",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 475000,
    "netPrice": 173420,
    "costPrice": 158400,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 365000,
        "netPrice": 133400,
        "costPrice": 121550
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 317000,
        "netPrice": 116000,
        "costPrice": 115200
      }
    ]
  },
  {
    "id": "REN-03-017",
    "categoryId": "cat-2",
    "category": "VCB・ｺﾝﾀｸﾀ",
    "name": "VCS 6.6kV 400A 常時励磁 ACDC100V",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 683000,
    "netPrice": 289380,
    "costPrice": 248400,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 525000,
        "netPrice": 222600,
        "costPrice": 222600
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 456000,
        "netPrice": 166000,
        "costPrice": 165600
      }
    ]
  },
  {
    "id": "REN-03-020",
    "categoryId": "cat-2",
    "category": "VCB・ｺﾝﾀｸﾀ",
    "name": "ｺﾝﾋﾞﾈｰｼｮﾝﾕﾆｯﾄ 6.6kV 200A 常時励磁 ACDC100V 引出形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1040000,
    "netPrice": 347600,
    "costPrice": 347600,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 926000,
        "netPrice": 309050,
        "costPrice": 309050,
        "sourceName": "ｺﾝﾋﾞﾈｰｼｮﾝﾕﾆｯﾄ 6.6kV 200A 常時励磁 DC100V 引出形"
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 695000,
        "netPrice": 233000,
        "costPrice": 233000,
        "sourceName": "ｺﾝﾋﾞﾈｰｼｮﾝﾕﾆｯﾄ 6.6kV 200A 常時励磁 DC100V 引出形"
      }
    ]
  },
  {
    "id": "REN-03-021",
    "categoryId": "cat-2",
    "category": "VCB・ｺﾝﾀｸﾀ",
    "name": "ｺﾝﾋﾞﾈｰｼｮﾝﾕﾆｯﾄ 6.6kV 200A 常時励磁 ACDC100V 薄型盤用",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1040000,
    "netPrice": 347490,
    "costPrice": 347600,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 800000,
        "netPrice": 267300,
        "costPrice": 267300,
        "sourceName": "ｺﾝﾋﾞﾈｰｼｮﾝﾕﾆｯﾄ 6.6kV 200A 常時励磁 ACDC100V 固定形"
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 586000,
        "netPrice": 233000,
        "costPrice": 233000,
        "sourceName": "ｺﾝﾋﾞﾈｰｼｮﾝﾕﾆｯﾄ 6.6kV 200A 常時励磁 ACDC100V 固定形"
      }
    ]
  },
  {
    "id": "REN-03-024",
    "categoryId": "cat-2",
    "category": "VCB・ｺﾝﾀｸﾀ",
    "name": "VSS 6.6kV 400A DC100V (分割引出型)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1874000,
    "netPrice": 937000,
    "costPrice": 937000,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 1874000,
        "netPrice": 937000,
        "costPrice": 937000
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 1874000,
        "netPrice": 937000,
        "costPrice": 937000
      }
    ]
  },
  {
    "id": "REN-03-025",
    "categoryId": "cat-2",
    "category": "VCB・ｺﾝﾀｸﾀ",
    "name": "VSS 6.6kV 600A DC100V (分割引出型)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1968000,
    "netPrice": 984000,
    "costPrice": 984000,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 1968000,
        "netPrice": 984000,
        "costPrice": 984000
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 1968000,
        "netPrice": 984000,
        "costPrice": 984000
      }
    ]
  },
  {
    "id": "REN-03-026",
    "categoryId": "cat-2",
    "category": "VCB・ｺﾝﾀｸﾀ",
    "name": "VSS 6.6kV 400A AC100V (分割引出型)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1894000,
    "netPrice": 947000,
    "costPrice": 947000,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 1894000,
        "netPrice": 947000,
        "costPrice": 947000
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 1894000,
        "netPrice": 947000,
        "costPrice": 947000
      }
    ]
  },
  {
    "id": "REN-03-027",
    "categoryId": "cat-2",
    "category": "VCB・ｺﾝﾀｸﾀ",
    "name": "VSS 6.6kV 600A AC100V (分割引出型)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1990000,
    "netPrice": 995000,
    "costPrice": 995000,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 1990000,
        "netPrice": 995000,
        "costPrice": 995000
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 1990000,
        "netPrice": 995000,
        "costPrice": 995000
      }
    ]
  },
  {
    "id": "REN-03-036",
    "categoryId": "cat-2",
    "category": "VCB・ｺﾝﾀｸﾀ",
    "name": "遠方手動操作器",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 139900,
    "netPrice": 77000,
    "costPrice": 52923,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 114000,
        "netPrice": 54000,
        "costPrice": 54000
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 114000,
        "netPrice": 54000,
        "costPrice": 54000
      }
    ]
  },
  {
    "id": "REN-04-005",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "OCR 電圧引外し",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 41000,
    "netPrice": 25000,
    "costPrice": 17920,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 41000,
        "netPrice": 25000,
        "costPrice": 17920
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 41200,
        "netPrice": 28800,
        "costPrice": 20600
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 35800,
        "netPrice": 25000,
        "costPrice": 17840
      }
    ]
  },
  {
    "id": "REN-04-006",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "OCR 電流引外し",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 41000,
    "netPrice": 25000,
    "costPrice": 17920,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 41000,
        "netPrice": 25000,
        "costPrice": 17920
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 41200,
        "netPrice": 28800,
        "costPrice": 20600
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 35800,
        "netPrice": 25000,
        "costPrice": 17840
      }
    ]
  },
  {
    "id": "REN-04-007",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "DGR 電圧引外し",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 137400,
    "netPrice": 50600,
    "costPrice": 49560,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 137400,
        "netPrice": 50600,
        "costPrice": 49560
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 137500,
        "netPrice": 50600,
        "costPrice": 49600
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 119500,
        "netPrice": 44000,
        "costPrice": 43110
      }
    ]
  },
  {
    "id": "REN-04-008",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "GR 電圧引外し",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 47100,
    "netPrice": 19600,
    "costPrice": 19220,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 47100,
        "netPrice": 19600,
        "costPrice": 19220
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 47200,
        "netPrice": 19600,
        "costPrice": 19300
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 41000,
        "netPrice": 17000,
        "costPrice": 16760
      }
    ]
  },
  {
    "id": "REN-04-009",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "GR 電流引外し",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 52000,
    "netPrice": 21900,
    "costPrice": 20920,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 52000,
        "netPrice": 21900,
        "costPrice": 20920
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 52100,
        "netPrice": 21900,
        "costPrice": 21000
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 45300,
        "netPrice": 19000,
        "costPrice": 18210
      }
    ]
  },
  {
    "id": "REN-04-010",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "UVR",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 67200,
    "netPrice": 34500,
    "costPrice": 29400,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 67200,
        "netPrice": 34500,
        "costPrice": 29400
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 67300,
        "netPrice": 34500,
        "costPrice": 33700
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 58500,
        "netPrice": 30000,
        "costPrice": 29280
      }
    ]
  },
  {
    "id": "REN-04-011",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "OVR",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 64200,
    "netPrice": 35000,
    "costPrice": 28070,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 64200,
        "netPrice": 35000,
        "costPrice": 28070
      }
    ]
  },
  {
    "id": "REN-04-013",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "ZPD",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 121800,
    "netPrice": 45000,
    "costPrice": 44140,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 121800,
        "netPrice": 45000,
        "costPrice": 44140
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 100100,
        "netPrice": 36800,
        "costPrice": 36300
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 87000,
        "netPrice": 32000,
        "costPrice": 31550
      }
    ]
  },
  {
    "id": "REN-04-015",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "ZCT 貫通形 53Φ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 18200,
    "netPrice": 7400,
    "costPrice": 7400,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 18200,
        "netPrice": 7400,
        "costPrice": 7400
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 15000,
        "netPrice": 6900,
        "costPrice": 6100
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 13000,
        "netPrice": 6000,
        "costPrice": 5260
      }
    ]
  },
  {
    "id": "REN-04-016",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "ZCT 貫通形 68Φ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 37100,
    "netPrice": 15000,
    "costPrice": 14940,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 37100,
        "netPrice": 15000,
        "costPrice": 14940
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 30500,
        "netPrice": 12700,
        "costPrice": 12400
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 26500,
        "netPrice": 11000,
        "costPrice": 10710
      }
    ]
  },
  {
    "id": "REN-04-017",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "ZCT 貫通形 90Φ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 94500,
    "netPrice": 45000,
    "costPrice": 44100,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 94500,
        "netPrice": 45000,
        "costPrice": 44100
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 77700,
        "netPrice": 41400,
        "costPrice": 41400
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 67500,
        "netPrice": 36000,
        "costPrice": 36000
      }
    ]
  },
  {
    "id": "REN-04-018",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "ZCT 貫通形 110Φ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 121800,
    "netPrice": 57000,
    "costPrice": 56840,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 121800,
        "netPrice": 57000,
        "costPrice": 56840
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 100100,
        "netPrice": 54100,
        "costPrice": 53400
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 87000,
        "netPrice": 47000,
        "costPrice": 46400
      }
    ]
  },
  {
    "id": "REN-04-019",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "ZCT 貫通形 160Φ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 228400,
    "netPrice": 95000,
    "costPrice": 94080,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 228400,
        "netPrice": 95000,
        "costPrice": 94080
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 187700,
        "netPrice": 88600,
        "costPrice": 88400
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 163200,
        "netPrice": 77000,
        "costPrice": 76800
      }
    ]
  },
  {
    "id": "REN-04-020",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "ZCT 貫通形 250Φ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 642600,
    "netPrice": 265000,
    "costPrice": 264600,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 642600,
        "netPrice": 265000,
        "costPrice": 264600
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 527900,
        "netPrice": 248400,
        "costPrice": 248400
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 459000,
        "netPrice": 216000,
        "costPrice": 216000
      }
    ]
  },
  {
    "id": "REN-04-021",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "ZCT 分割形 52Φ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 79500,
    "netPrice": 31000,
    "costPrice": 30070,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 79500,
        "netPrice": 31000,
        "costPrice": 30070
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 65400,
        "netPrice": 25300,
        "costPrice": 24800
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 56800,
        "netPrice": 22000,
        "costPrice": 21480
      }
    ]
  },
  {
    "id": "REN-04-022",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "ZCT 分割形 77Φ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 105200,
    "netPrice": 47000,
    "costPrice": 46060,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 105200,
        "netPrice": 47000,
        "costPrice": 46060
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 86500,
        "netPrice": 43700,
        "costPrice": 43300
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 75200,
        "netPrice": 38000,
        "costPrice": 37600
      }
    ]
  },
  {
    "id": "REN-04-023",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "ZCT 分割形 112Φ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 152300,
    "netPrice": 67000,
    "costPrice": 66640,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 152300,
        "netPrice": 67000,
        "costPrice": 66640
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 125200,
        "netPrice": 63300,
        "costPrice": 62600
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 108800,
        "netPrice": 55000,
        "costPrice": 54400
      }
    ]
  },
  {
    "id": "REN-04-025",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "電流トリップ補助箱",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 20500,
    "netPrice": 10400,
    "costPrice": 9590,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 20500,
        "netPrice": 10400,
        "costPrice": 9590
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 17000,
        "netPrice": 10400,
        "costPrice": 10200
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 14700,
        "netPrice": 9000,
        "costPrice": 8820
      }
    ]
  },
  {
    "id": "REN-04-026",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "CTD",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 62790,
    "netPrice": 36000,
    "costPrice": 35880,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 48300,
        "netPrice": 27600,
        "costPrice": 27600
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 48300,
        "netPrice": 27600,
        "costPrice": 27600
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 42000,
        "netPrice": 24000,
        "costPrice": 24000
      }
    ]
  },
  {
    "id": "REN-04-027",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "CTD (ｴﾈｾｰﾊﾞ用)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 70720,
    "netPrice": 44000,
    "costPrice": 43550,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 54400,
        "netPrice": 34000,
        "costPrice": 33500
      }
    ]
  },
  {
    "id": "REN-04-029",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "APFC",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 228000,
    "netPrice": 105000,
    "costPrice": 105000,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 228000,
        "netPrice": 105000,
        "costPrice": 105000
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 228000,
        "netPrice": 105000,
        "costPrice": 105000
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 228000,
        "netPrice": 105000,
        "costPrice": 105000
      }
    ]
  },
  {
    "id": "REN-04-032",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "OCR 電圧引外し",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 53900,
    "netPrice": 34800,
    "costPrice": 34800,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 41400,
        "netPrice": 29000,
        "costPrice": 29000
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 41400,
        "netPrice": 29000,
        "costPrice": 29000,
        "sourceName": "OCR 電圧引外し LOC-23"
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 41400,
        "netPrice": 29000,
        "costPrice": 29000,
        "sourceName": "OCR 電圧引外し LOC-23"
      }
    ]
  },
  {
    "id": "REN-04-033",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "OCR 電流引外し",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 53900,
    "netPrice": 34800,
    "costPrice": 34800,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 41400,
        "netPrice": 29000,
        "costPrice": 29000
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 41400,
        "netPrice": 29000,
        "costPrice": 29000,
        "sourceName": "OCR 電流引外し LOC-24"
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 41400,
        "netPrice": 29000,
        "costPrice": 29000,
        "sourceName": "OCR 電流引外し LOC-24"
      }
    ]
  },
  {
    "id": "REN-04-034",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "UVR",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 89700,
    "netPrice": 57700,
    "costPrice": 57700,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 74700,
        "netPrice": 53000,
        "costPrice": 52400
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 74750,
        "netPrice": 52000,
        "costPrice": 52000,
        "sourceName": "UVR LUV-2A"
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 74750,
        "netPrice": 52000,
        "costPrice": 52000,
        "sourceName": "UVR LUV-2A"
      }
    ]
  },
  {
    "id": "REN-04-036",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "DGR 電圧引き外し方式 (親機)",
    "spec": "★",
    "defaultQty": 1,
    "listPrice": 169000,
    "netPrice": 84000,
    "costPrice": 84000,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 130000,
        "netPrice": 70000,
        "costPrice": 70000
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 130000,
        "netPrice": 70000,
        "costPrice": 70000,
        "sourceName": "LDG-73 電圧引き外し方式"
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 130000,
        "netPrice": 70000,
        "costPrice": 70000,
        "sourceName": "LDG-73 電圧引き外し方式"
      }
    ]
  },
  {
    "id": "REN-04-037",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "DGR 電流引き外し方式 (親機)",
    "spec": "★",
    "defaultQty": 1,
    "listPrice": 188500,
    "netPrice": 92400,
    "costPrice": 92400,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 145000,
        "netPrice": 77000,
        "costPrice": 77000
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 145000,
        "netPrice": 77000,
        "costPrice": 77000,
        "sourceName": "LDG-73K 電流引き外し方式"
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 145000,
        "netPrice": 77000,
        "costPrice": 77000,
        "sourceName": "LDG-73K 電流引き外し方式"
      }
    ]
  },
  {
    "id": "REN-04-038",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "DGR ﾃﾞｼﾞﾀﾙ形 電圧引き外し方式 (親機)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 172500,
    "netPrice": 81000,
    "costPrice": 85100,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 150000,
        "netPrice": 81000,
        "costPrice": 81000
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 150000,
        "netPrice": 81000,
        "costPrice": 81000,
        "sourceName": "LDG-83 デジタル形、電圧引き外し方式"
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 150000,
        "netPrice": 81000,
        "costPrice": 81000,
        "sourceName": "LDG-83 デジタル形、電圧引き外し方式"
      }
    ]
  },
  {
    "id": "REN-04-040",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "DGR 電圧引き外し方式 (子機)",
    "spec": "★",
    "defaultQty": 1,
    "listPrice": 156000,
    "netPrice": 81600,
    "costPrice": 81600,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 120000,
        "netPrice": 68000,
        "costPrice": 68000
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 120000,
        "netPrice": 68000,
        "costPrice": 68000,
        "sourceName": "LDG-71 電圧引き外し方式"
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 120000,
        "netPrice": 68000,
        "costPrice": 68000,
        "sourceName": "LDG-71 電圧引き外し方式"
      }
    ]
  },
  {
    "id": "REN-04-041",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "DGR 電流引き外し方式 (子機)",
    "spec": "★",
    "defaultQty": 1,
    "listPrice": 175500,
    "netPrice": 90000,
    "costPrice": 90000,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 135000,
        "netPrice": 75000,
        "costPrice": 75000
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 135000,
        "netPrice": 75000,
        "costPrice": 75000,
        "sourceName": "LDG-71K 電流引き外し方式"
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 135000,
        "netPrice": 75000,
        "costPrice": 75000,
        "sourceName": "LDG-71K 電流引き外し方式"
      }
    ]
  },
  {
    "id": "REN-04-042",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "DGR ﾃﾞｼﾞﾀﾙ形 電圧引き外し方式 (子機)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 161000,
    "netPrice": 78000,
    "costPrice": 81900,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 140000,
        "netPrice": 78000,
        "costPrice": 78000
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 140000,
        "netPrice": 78000,
        "costPrice": 78000,
        "sourceName": "LDG-81 デジタル形、電圧引き外し方式"
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 140000,
        "netPrice": 78000,
        "costPrice": 78000,
        "sourceName": "LDG-81 デジタル形、電圧引き外し方式"
      }
    ]
  },
  {
    "id": "REN-04-044",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "GR 手動復帰 露出形",
    "spec": "★",
    "defaultQty": 1,
    "listPrice": 29900,
    "netPrice": 17300,
    "costPrice": 17300,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 23000,
        "netPrice": 15000,
        "costPrice": 14400
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 23000,
        "netPrice": 15000,
        "costPrice": 14400
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 23000,
        "netPrice": 15000,
        "costPrice": 14400
      }
    ]
  },
  {
    "id": "REN-04-045",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "GR 手動復帰 埋込形",
    "spec": "★",
    "defaultQty": 1,
    "listPrice": 31200,
    "netPrice": 18800,
    "costPrice": 18800,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 24000,
        "netPrice": 16000,
        "costPrice": 15600
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 24000,
        "netPrice": 16000,
        "costPrice": 15600
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 24000,
        "netPrice": 16000,
        "costPrice": 15600
      }
    ]
  },
  {
    "id": "REN-04-046",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "GR 手動/自動復帰 埋込形",
    "spec": "★",
    "defaultQty": 1,
    "listPrice": 29900,
    "netPrice": 17800,
    "costPrice": 17800,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 23000,
        "netPrice": 15000,
        "costPrice": 14800
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 23000,
        "netPrice": 15000,
        "costPrice": 14800
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 23000,
        "netPrice": 15000,
        "costPrice": 14800
      }
    ]
  },
  {
    "id": "REN-04-047",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "GR 手動/自動復帰 埋込形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 45900,
    "netPrice": 26700,
    "costPrice": 26700,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 35300,
        "netPrice": 23000,
        "costPrice": 22200
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 35300,
        "netPrice": 23000,
        "costPrice": 22200
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 35300,
        "netPrice": 23000,
        "costPrice": 22200
      }
    ]
  },
  {
    "id": "REN-04-049",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "ZPD 碍子形",
    "spec": "★",
    "defaultQty": 1,
    "listPrice": 123600,
    "netPrice": 80000,
    "costPrice": 79200,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 103000,
        "netPrice": 72000,
        "costPrice": 72000
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 68400,
        "netPrice": 48000,
        "costPrice": 48000,
        "sourceName": "ZPD ZPC-9B 碍子形"
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 68400,
        "netPrice": 48000,
        "costPrice": 48000,
        "sourceName": "ZPD ZPC-9B 碍子形"
      }
    ]
  },
  {
    "id": "REN-04-050",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "補助継電器",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 23400,
    "netPrice": 15200,
    "costPrice": 15200,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 18000,
        "netPrice": 12600,
        "costPrice": 12600
      }
    ]
  },
  {
    "id": "REN-04-053",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "ZCT 貫通形 41Φ 100A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 20500,
    "netPrice": 15000,
    "costPrice": 14400,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 20500,
        "netPrice": 15000,
        "costPrice": 14400
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 18500,
        "netPrice": 13000,
        "costPrice": 12900,
        "sourceName": "M64 64Φ 200A"
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 18500,
        "netPrice": 13000,
        "costPrice": 12900,
        "sourceName": "M64 64Φ 200A"
      }
    ]
  },
  {
    "id": "REN-04-054",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "ZCT 貫通形 64Φ 200A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 23200,
    "netPrice": 17000,
    "costPrice": 16300,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 23200,
        "netPrice": 17000,
        "costPrice": 16300
      }
    ]
  },
  {
    "id": "REN-04-055",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "ZCT 貫通形 106Φ 400A or 600A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 40500,
    "netPrice": 29000,
    "costPrice": 28400,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 40500,
        "netPrice": 29000,
        "costPrice": 28400
      }
    ]
  },
  {
    "id": "REN-04-056",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "ZCT 貫通形 120Φ 800A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 77800,
    "netPrice": 55000,
    "costPrice": 54500,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 77800,
        "netPrice": 55000,
        "costPrice": null
      }
    ]
  },
  {
    "id": "REN-04-057",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "ZCT 貫通形 156Φ 1000A (受注品)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 99500,
    "netPrice": 70000,
    "costPrice": 69700,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 99500,
        "netPrice": 70000,
        "costPrice": 69700
      }
    ]
  },
  {
    "id": "REN-04-058",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "ZCT 貫通形 240Φ 2200A (受注品)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 256100,
    "netPrice": 180000,
    "costPrice": 179300,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 256100,
        "netPrice": 180000,
        "costPrice": 179300
      }
    ]
  },
  {
    "id": "REN-04-060",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "ZCT 分割形 55Φ 300A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 65800,
    "netPrice": 47000,
    "costPrice": 46100,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 65800,
        "netPrice": 47000,
        "costPrice": 46100
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 101400,
        "netPrice": 71000,
        "costPrice": 71000,
        "sourceName": "DM70 70Φ 400A"
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 101400,
        "netPrice": 71000,
        "costPrice": 71000,
        "sourceName": "DM70 70Φ 400A"
      }
    ]
  },
  {
    "id": "REN-04-061",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "ZCT 分割形 70Φ 400A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 111600,
    "netPrice": 79000,
    "costPrice": 78200,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 111600,
        "netPrice": 79000,
        "costPrice": 78200
      }
    ]
  },
  {
    "id": "REN-04-062",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "ZCT 分割形 100Φ 600A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 158800,
    "netPrice": 112000,
    "costPrice": 111200,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 158800,
        "netPrice": 112000,
        "costPrice": 111200
      }
    ]
  },
  {
    "id": "REN-04-064",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "CF-123 M41用",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1500,
    "netPrice": 1000,
    "costPrice": 700,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": null,
        "netPrice": 1000,
        "costPrice": null
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 1800,
        "netPrice": 1200,
        "costPrice": 1200,
        "sourceName": "CF-124 M64用"
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 1800,
        "netPrice": 1200,
        "costPrice": 1200,
        "sourceName": "CF-124 M64用"
      }
    ]
  },
  {
    "id": "REN-04-065",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "CF-124 M64用",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2400,
    "netPrice": 2000,
    "costPrice": 1600,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 1800,
        "netPrice": 2000,
        "costPrice": 1200
      }
    ]
  },
  {
    "id": "REN-04-066",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "CF-125 M106 (400A) 用",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 5500,
    "netPrice": 4000,
    "costPrice": 3400,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": null,
        "netPrice": 3000,
        "costPrice": null
      }
    ]
  },
  {
    "id": "REN-04-067",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "CF-126 M106 (600A) 用",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 5500,
    "netPrice": 4000,
    "costPrice": 3400,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": null,
        "netPrice": 3000,
        "costPrice": null
      }
    ]
  },
  {
    "id": "REN-04-071",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "DGR",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 113000,
    "netPrice": 68000,
    "costPrice": 67670,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 113000,
        "netPrice": 68000,
        "costPrice": 67670
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 79500,
        "netPrice": 57000,
        "costPrice": 56320,
        "sourceName": "ZPD"
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 79500,
        "netPrice": 57000,
        "costPrice": 56320,
        "sourceName": "ZPD"
      }
    ]
  },
  {
    "id": "REN-04-072",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "補助電源装置",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 24000,
    "netPrice": 17000,
    "costPrice": 16110,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 24000,
        "netPrice": 17000,
        "costPrice": 16110
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 95500,
        "netPrice": 67000,
        "costPrice": 66440,
        "sourceName": "ZPD ｹｰﾌﾞﾙ3m"
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 95500,
        "netPrice": 67000,
        "costPrice": 66440,
        "sourceName": "ZPD ｹｰﾌﾞﾙ3m"
      }
    ]
  },
  {
    "id": "REN-04-073",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "ZPD",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 79500,
    "netPrice": 57000,
    "costPrice": 56320,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 79500,
        "netPrice": 57000,
        "costPrice": 56320
      },
      {
        "label": "三菱15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 97500,
        "netPrice": 70000,
        "costPrice": 69780,
        "sourceName": "ZCT 400A (分割形)"
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 97500,
        "netPrice": 70000,
        "costPrice": 69780,
        "sourceName": "ZCT 400A (分割形)"
      }
    ]
  },
  {
    "id": "REN-04-074",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "ZPD ｹｰﾌﾞﾙ3m",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 95500,
    "netPrice": 67000,
    "costPrice": 66440,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 95500,
        "netPrice": 67000,
        "costPrice": 66440
      }
    ]
  },
  {
    "id": "REN-04-075",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "ZCT 400A (分割形)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 97500,
    "netPrice": 70000,
    "costPrice": 69780,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 97500,
        "netPrice": 70000,
        "costPrice": 69780
      }
    ]
  },
  {
    "id": "REN-04-076",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "OCR 電圧引外し 電流引外し",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 35500,
    "netPrice": null,
    "costPrice": 24890,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 35500,
        "netPrice": null,
        "costPrice": 24890
      }
    ]
  },
  {
    "id": "REN-04-079",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "DRPR-72",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 183600,
    "netPrice": 141370,
    "costPrice": 141370,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 183600,
        "netPrice": 141370,
        "costPrice": 141370
      }
    ]
  },
  {
    "id": "REN-04-080",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "上記用ｸﾗﾝﾌﾟCT",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 2800,
    "netPrice": 2160,
    "costPrice": 2160,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 2800,
        "netPrice": 2160,
        "costPrice": 2160
      }
    ]
  },
  {
    "id": "REN-04-083",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "パルスピック PC-11B",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 200000,
    "netPrice": 115000,
    "costPrice": 114080,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "保護継電器15%up",
        "effectiveDate": null,
        "listPrice": 200000,
        "netPrice": 115000,
        "costPrice": 114080
      }
    ]
  },
  {
    "id": "REN-04-087",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "GR 電流・電圧引き外し 埋込型",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 51800,
    "netPrice": 20800,
    "costPrice": 20800,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-04-088",
    "categoryId": "cat-3",
    "category": "継電器",
    "name": "ZCT 分割形 400A Φ72",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 146800,
    "netPrice": 58900,
    "costPrice": 58900,
    "priceLabel": "三菱CTD 30%up・光商工 R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-05-005",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "ME110GR-MBH",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 141000,
    "netPrice": 65000,
    "costPrice": 64400,
    "priceLabel": "現行",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-05-008",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "OVGR・RPR",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 150500,
    "netPrice": 78000,
    "costPrice": 79700,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": [
      {
        "label": "太陽光",
        "effectiveDate": null,
        "listPrice": 130800,
        "netPrice": 80000,
        "costPrice": 77000
      }
    ]
  },
  {
    "id": "REN-05-009",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "ZPD",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 100100,
    "netPrice": 37000,
    "costPrice": 36300,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": [
      {
        "label": "太陽光",
        "effectiveDate": null,
        "listPrice": 121800,
        "netPrice": 45000,
        "costPrice": 44140
      }
    ]
  },
  {
    "id": "REN-05-010",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "ブロック電源",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": []
  },
  {
    "id": "REN-05-011",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "瞬停対策ブロック",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": []
  },
  {
    "id": "REN-05-012",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "バスラインコネクタ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": []
  },
  {
    "id": "REN-05-013",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "ZVTT",
    "spec": "",
    "defaultQty": 2,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": []
  },
  {
    "id": "REN-05-016",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "RPR",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 91100,
    "netPrice": 48000,
    "costPrice": 48000,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": [
      {
        "label": "太陽光",
        "effectiveDate": null,
        "listPrice": 79200,
        "netPrice": 47000,
        "costPrice": 46300,
        "sourceName": "OVGR"
      }
    ]
  },
  {
    "id": "REN-05-017",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "ZPD",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 100100,
    "netPrice": 37000,
    "costPrice": 36300,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": [
      {
        "label": "太陽光",
        "effectiveDate": null,
        "listPrice": 121800,
        "netPrice": 45000,
        "costPrice": 44140
      }
    ]
  },
  {
    "id": "REN-05-018",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "ブロック電源",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": []
  },
  {
    "id": "REN-05-019",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "瞬停対策ブロック",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": []
  },
  {
    "id": "REN-05-020",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "バスラインコネクタ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": []
  },
  {
    "id": "REN-05-021",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "ZVTT",
    "spec": "",
    "defaultQty": 2,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": []
  },
  {
    "id": "REN-05-024",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "OVGR",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 65000,
    "netPrice": 49000,
    "costPrice": 49000,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": [
      {
        "label": "太陽光",
        "effectiveDate": null,
        "listPrice": 54100,
        "netPrice": 41000,
        "costPrice": 40800
      }
    ]
  },
  {
    "id": "REN-05-025",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "RPR",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 66000,
    "netPrice": 50000,
    "costPrice": 50000,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": [
      {
        "label": "太陽光",
        "effectiveDate": null,
        "listPrice": 54700,
        "netPrice": 42000,
        "costPrice": 41500
      }
    ]
  },
  {
    "id": "REN-05-026",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "2ユニットケース",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 23000,
    "netPrice": 18000,
    "costPrice": 18000,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": [
      {
        "label": "太陽光",
        "effectiveDate": null,
        "listPrice": 19000,
        "netPrice": 15000,
        "costPrice": 14300
      }
    ]
  },
  {
    "id": "REN-05-027",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "ZPD",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 79500,
    "netPrice": 58000,
    "costPrice": 58000,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": [
      {
        "label": "太陽光",
        "effectiveDate": null,
        "listPrice": 66200,
        "netPrice": 49000,
        "costPrice": 48300
      }
    ]
  },
  {
    "id": "REN-05-028",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "ブロック電源",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 15500,
    "netPrice": 12000,
    "costPrice": 12000,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": [
      {
        "label": "太陽光",
        "effectiveDate": null,
        "listPrice": 14100,
        "netPrice": 11000,
        "costPrice": 10520
      }
    ]
  },
  {
    "id": "REN-05-029",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "瞬停対策ブロック",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 25300,
    "netPrice": 20000,
    "costPrice": 20000,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": [
      {
        "label": "太陽光",
        "effectiveDate": null,
        "listPrice": 23000,
        "netPrice": 18000,
        "costPrice": 17820
      }
    ]
  },
  {
    "id": "REN-05-030",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "バスラインコネクタ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 640,
    "netPrice": 640,
    "costPrice": 640,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": [
      {
        "label": "太陽光",
        "effectiveDate": null,
        "listPrice": 580,
        "netPrice": 600,
        "costPrice": 550
      }
    ]
  },
  {
    "id": "REN-05-031",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "ZVTT",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 1000,
    "netPrice": 600,
    "costPrice": 600,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": [
      {
        "label": "太陽光",
        "effectiveDate": null,
        "listPrice": 2000,
        "netPrice": 1000,
        "costPrice": 1000
      }
    ]
  },
  {
    "id": "REN-05-035",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "OVGR",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 65000,
    "netPrice": 49000,
    "costPrice": 49000,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": [
      {
        "label": "太陽光",
        "effectiveDate": null,
        "listPrice": 54100,
        "netPrice": 41000,
        "costPrice": 40800
      }
    ]
  },
  {
    "id": "REN-05-036",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "フリーユニット",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 14100,
    "netPrice": 11000,
    "costPrice": 11000,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": [
      {
        "label": "太陽光",
        "effectiveDate": null,
        "listPrice": 11700,
        "netPrice": 9000,
        "costPrice": 8800
      }
    ]
  },
  {
    "id": "REN-05-037",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "2ユニットケース",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 23000,
    "netPrice": 18000,
    "costPrice": 18000,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": [
      {
        "label": "太陽光",
        "effectiveDate": null,
        "listPrice": 19000,
        "netPrice": 15000,
        "costPrice": 14300
      }
    ]
  },
  {
    "id": "REN-05-038",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "ZPD",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 79500,
    "netPrice": 58000,
    "costPrice": 58000,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": [
      {
        "label": "太陽光",
        "effectiveDate": null,
        "listPrice": 66200,
        "netPrice": 49000,
        "costPrice": 48300
      }
    ]
  },
  {
    "id": "REN-05-039",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "ブロック電源",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 15500,
    "netPrice": 12000,
    "costPrice": 12000,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": [
      {
        "label": "太陽光",
        "effectiveDate": null,
        "listPrice": 14100,
        "netPrice": 11000,
        "costPrice": 10520
      }
    ]
  },
  {
    "id": "REN-05-040",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "瞬停対策ブロック",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 25300,
    "netPrice": 20000,
    "costPrice": 20000,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": [
      {
        "label": "太陽光",
        "effectiveDate": null,
        "listPrice": 23000,
        "netPrice": 18000,
        "costPrice": 17820
      }
    ]
  },
  {
    "id": "REN-05-041",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "バスラインコネクタ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 640,
    "netPrice": 640,
    "costPrice": 640,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": [
      {
        "label": "太陽光",
        "effectiveDate": null,
        "listPrice": 580,
        "netPrice": 600,
        "costPrice": 550
      }
    ]
  },
  {
    "id": "REN-05-042",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "ZVTT",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 1000,
    "netPrice": 600,
    "costPrice": 600,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": [
      {
        "label": "太陽光",
        "effectiveDate": null,
        "listPrice": 2000,
        "netPrice": 1000,
        "costPrice": 1000
      }
    ]
  },
  {
    "id": "REN-05-045",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "OVGR",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": 50000,
    "costPrice": 50000,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": [
      {
        "label": "太陽光",
        "effectiveDate": null,
        "listPrice": null,
        "netPrice": 50000,
        "costPrice": 50000
      }
    ]
  },
  {
    "id": "REN-05-046",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "ZPD ZPC-9B 碍子形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 68400,
    "netPrice": 48000,
    "costPrice": 48000,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": [
      {
        "label": "太陽光",
        "effectiveDate": null,
        "listPrice": 123600,
        "netPrice": 80000,
        "costPrice": 79200,
        "sourceName": "ZPD 碍子形"
      }
    ]
  },
  {
    "id": "REN-05-050",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "RS485回路用SPD",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 20600,
    "netPrice": 16000,
    "costPrice": 15040,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": [
      {
        "label": "太陽光",
        "effectiveDate": null,
        "listPrice": 123600,
        "netPrice": 80000,
        "costPrice": 79200,
        "sourceName": "ZPD 碍子形"
      }
    ]
  },
  {
    "id": "REN-05-055",
    "categoryId": "cat-4",
    "category": "太陽光",
    "name": "RS485回路用SPD",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 20600,
    "netPrice": 16000,
    "costPrice": 15040,
    "priceLabel": "各々UP R6.10.1",
    "effectiveDate": "2024-10-01",
    "history": [
      {
        "label": "太陽光",
        "effectiveDate": null,
        "listPrice": 123600,
        "netPrice": 80000,
        "costPrice": 79200,
        "sourceName": "ZPD 碍子形"
      }
    ]
  },
  {
    "id": "REN-06-006",
    "categoryId": "cat-5",
    "category": "ﾏﾙﾁ",
    "name": "ME110GR",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 105000,
    "netPrice": 42000,
    "costPrice": 41860,
    "priceLabel": "現行",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-06-007",
    "categoryId": "cat-5",
    "category": "ﾏﾙﾁ",
    "name": "ME110GR-4AP",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 117000,
    "netPrice": 47000,
    "costPrice": 46690,
    "priceLabel": "現行",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-06-008",
    "categoryId": "cat-5",
    "category": "ﾏﾙﾁ",
    "name": "ME110GR-4APH",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 121000,
    "netPrice": 49000,
    "costPrice": 48300,
    "priceLabel": "現行",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-06-009",
    "categoryId": "cat-5",
    "category": "ﾏﾙﾁ",
    "name": "ME110GR-4A2P",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 121000,
    "netPrice": 49000,
    "costPrice": 48300,
    "priceLabel": "現行",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-06-010",
    "categoryId": "cat-5",
    "category": "ﾏﾙﾁ",
    "name": "ME110GR-6APH",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 141000,
    "netPrice": 57000,
    "costPrice": 56350,
    "priceLabel": "現行",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-06-011",
    "categoryId": "cat-5",
    "category": "ﾏﾙﾁ",
    "name": "ME110GR-6A2P",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 141000,
    "netPrice": 57000,
    "costPrice": 56350,
    "priceLabel": "現行",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-06-012",
    "categoryId": "cat-5",
    "category": "ﾏﾙﾁ",
    "name": "ME110GF",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 90700,
    "netPrice": 37000,
    "costPrice": 36260,
    "priceLabel": "現行",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-06-013",
    "categoryId": "cat-5",
    "category": "ﾏﾙﾁ",
    "name": "ME110GF-2APH",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 103000,
    "netPrice": 42000,
    "costPrice": 41090,
    "priceLabel": "現行",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-06-014",
    "categoryId": "cat-5",
    "category": "ﾏﾙﾁ",
    "name": "ME110GF-PH",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 98800,
    "netPrice": 40000,
    "costPrice": 39480,
    "priceLabel": "現行",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-06-015",
    "categoryId": "cat-5",
    "category": "ﾏﾙﾁ",
    "name": "ME110GBA",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 80500,
    "netPrice": 33000,
    "costPrice": 32200,
    "priceLabel": "現行",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-06-016",
    "categoryId": "cat-5",
    "category": "ﾏﾙﾁ",
    "name": "ME110GBA-H",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 90700,
    "netPrice": 37000,
    "costPrice": 36260,
    "priceLabel": "現行",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-06-017",
    "categoryId": "cat-5",
    "category": "ﾏﾙﾁ",
    "name": "ME110GR-BH",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 174000,
    "netPrice": 70000,
    "costPrice": 69230,
    "priceLabel": "現行",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-06-018",
    "categoryId": "cat-5",
    "category": "ﾏﾙﾁ",
    "name": "ME110GR-CH",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 199000,
    "netPrice": 80000,
    "costPrice": 79800,
    "priceLabel": "現行",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-06-019",
    "categoryId": "cat-5",
    "category": "ﾏﾙﾁ",
    "name": "ME110GR-MBH",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 145000,
    "netPrice": 58000,
    "costPrice": 57960,
    "priceLabel": "現行",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-07-005",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "LBS 7.2kV 200A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 86800,
    "netPrice": 47740,
    "costPrice": 23280,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "三菱15%up",
        "effectiveDate": null,
        "listPrice": 66800,
        "netPrice": 36800,
        "costPrice": 17910
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 58100,
        "netPrice": 32000,
        "costPrice": 15590
      }
    ]
  },
  {
    "id": "REN-07-006",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "LBS 7.2kV 200A (G75A,G100A)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 114500,
    "netPrice": 62975,
    "costPrice": 31020,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "三菱15%up",
        "effectiveDate": null,
        "listPrice": 88100,
        "netPrice": 48500,
        "costPrice": 23880
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 76600,
        "netPrice": 43000,
        "costPrice": 26520
      }
    ]
  },
  {
    "id": "REN-07-007",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "LBS 7.2kV 200A (G80,G100,T88A)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 124500,
    "netPrice": 68475,
    "costPrice": 36610,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "三菱15%up",
        "effectiveDate": null,
        "listPrice": 95800,
        "netPrice": 52700,
        "costPrice": 28150
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 83300,
        "netPrice": 46000,
        "costPrice": 25010
      }
    ]
  },
  {
    "id": "REN-07-008",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "LBS 7.2kV 200A TC付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 120500,
    "netPrice": 66275,
    "costPrice": 34080,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "三菱15%up",
        "effectiveDate": null,
        "listPrice": 92700,
        "netPrice": 51800,
        "costPrice": 26200
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 80600,
        "netPrice": 45000,
        "costPrice": 22780
      }
    ]
  },
  {
    "id": "REN-07-009",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "LBS 7.2kV 200A TC・ﾋｭｰｽﾞ動作表示ｽｲｯﾁ付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 142600,
    "netPrice": 78430,
    "costPrice": 41880,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "三菱15%up",
        "effectiveDate": null,
        "listPrice": 109700,
        "netPrice": 61000,
        "costPrice": 32200
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 95400,
        "netPrice": 53000,
        "costPrice": 28490
      }
    ]
  },
  {
    "id": "REN-07-011",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "LBS 7.2kV 200A 絶縁ﾊﾞﾘｱ付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 124400,
    "netPrice": 68420,
    "costPrice": 35980,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "三菱15%up",
        "effectiveDate": null,
        "listPrice": 148600,
        "netPrice": 81800,
        "costPrice": 41830
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 12600,
        "netPrice": 6000,
        "costPrice": 5400,
        "sourceName": "絶縁バリア"
      }
    ]
  },
  {
    "id": "REN-07-012",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "絶縁バリア",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 18800,
    "netPrice": 10340,
    "costPrice": 6350,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "三菱15%up",
        "effectiveDate": null,
        "listPrice": 14500,
        "netPrice": 6900,
        "costPrice": 6300
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 12000,
        "netPrice": 6600,
        "costPrice": 3310,
        "sourceName": "ﾄﾘｯﾌﾟｺｲﾙ"
      }
    ]
  },
  {
    "id": "REN-07-013",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "ﾄﾘｯﾌﾟｺｲﾙ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 15800,
    "netPrice": 8690,
    "costPrice": 4640,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "三菱15%up",
        "effectiveDate": null,
        "listPrice": 13800,
        "netPrice": 7600,
        "costPrice": 3900
      }
    ]
  },
  {
    "id": "REN-07-015",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "LBS 7.2kV 200A (ﾋｭｰｽﾞ無し)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 108600,
    "netPrice": 59730,
    "costPrice": 31910,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "三菱15%up",
        "effectiveDate": null,
        "listPrice": 83500,
        "netPrice": 46000,
        "costPrice": 23000
      }
    ]
  },
  {
    "id": "REN-07-019",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "LBS 7.2kV 200A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 147300,
    "netPrice": 81015,
    "costPrice": 43240,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "三菱15%up",
        "effectiveDate": null,
        "listPrice": 113300,
        "netPrice": 63000,
        "costPrice": 33280
      }
    ]
  },
  {
    "id": "REN-07-020",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "LBS 7.2kV 200A (G150,G200A)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 134600,
    "netPrice": 74030,
    "costPrice": 41690,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "三菱15%up",
        "effectiveDate": null,
        "listPrice": 103500,
        "netPrice": 57000,
        "costPrice": 32050
      }
    ]
  },
  {
    "id": "REN-07-024",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "LBS用PF 7.2kV G5A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 8370,
    "netPrice": 4604,
    "costPrice": 2550,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "三菱15%up",
        "effectiveDate": null,
        "listPrice": 6440,
        "netPrice": 3600,
        "costPrice": 1960
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 5600,
        "netPrice": 3000,
        "costPrice": 1710
      }
    ]
  },
  {
    "id": "REN-07-025",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "LBS用PF 7.2kV G10A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 8370,
    "netPrice": 4604,
    "costPrice": 2550,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "三菱15%up",
        "effectiveDate": null,
        "listPrice": 6440,
        "netPrice": 3600,
        "costPrice": 1960
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 5600,
        "netPrice": 3000,
        "costPrice": 1710
      }
    ]
  },
  {
    "id": "REN-07-026",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "LBS用PF 7.2kV G20A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 9270,
    "netPrice": 5099,
    "costPrice": 3170,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "三菱15%up",
        "effectiveDate": null,
        "listPrice": 7130,
        "netPrice": 4000,
        "costPrice": 2440
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 6200,
        "netPrice": 4000,
        "costPrice": 2120
      }
    ]
  },
  {
    "id": "REN-07-027",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "LBS用PF 7.2kV G30A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 14350,
    "netPrice": 7893,
    "costPrice": 3930,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "三菱15%up",
        "effectiveDate": null,
        "listPrice": 11040,
        "netPrice": 6100,
        "costPrice": 3020
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 9600,
        "netPrice": 6000,
        "costPrice": 2620
      }
    ]
  },
  {
    "id": "REN-07-028",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "LBS用PF 7.2kV G40A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 17300,
    "netPrice": 9515,
    "costPrice": 4920,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "三菱15%up",
        "effectiveDate": null,
        "listPrice": 13340,
        "netPrice": 7400,
        "costPrice": 3790
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 11600,
        "netPrice": 7000,
        "costPrice": 3290
      }
    ]
  },
  {
    "id": "REN-07-029",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "LBS用PF 7.2kV G50A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 19600,
    "netPrice": 10780,
    "costPrice": 5720,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "三菱15%up",
        "effectiveDate": null,
        "listPrice": 15070,
        "netPrice": 8300,
        "costPrice": 4400
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 13100,
        "netPrice": 8000,
        "costPrice": 3830
      }
    ]
  },
  {
    "id": "REN-07-030",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "LBS用PF 7.2kV G60A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 22100,
    "netPrice": 12155,
    "costPrice": 6530,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "三菱15%up",
        "effectiveDate": null,
        "listPrice": 17000,
        "netPrice": 9400,
        "costPrice": 5020
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 14800,
        "netPrice": 9000,
        "costPrice": 4370
      }
    ]
  },
  {
    "id": "REN-07-031",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "LBS用PF 7.2kV G75A (T50A)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 29900,
    "netPrice": 16445,
    "costPrice": 8680,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "三菱15%up",
        "effectiveDate": null,
        "listPrice": 23000,
        "netPrice": 12700,
        "costPrice": 6680
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 20000,
        "netPrice": 11000,
        "costPrice": 5800
      }
    ]
  },
  {
    "id": "REN-07-032",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "LBS用PF 7.2kV G75A (T60A)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 29900,
    "netPrice": 16445,
    "costPrice": 8680,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "三菱15%up",
        "effectiveDate": null,
        "listPrice": 23000,
        "netPrice": 12700,
        "costPrice": 6680
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 20000,
        "netPrice": 11000,
        "costPrice": 5800
      }
    ]
  },
  {
    "id": "REN-07-033",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "LBS用PF 7.2kV G80A (T66A)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 49900,
    "netPrice": 27445,
    "costPrice": 21360,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "三菱15%up",
        "effectiveDate": null,
        "listPrice": 38400,
        "netPrice": 21200,
        "costPrice": 16440
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 33400,
        "netPrice": 19000,
        "costPrice": 14280
      }
    ]
  },
  {
    "id": "REN-07-034",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "LBS用PF 7.2kV G100A (T76A)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 55400,
    "netPrice": 30470,
    "costPrice": 23700,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "三菱15%up",
        "effectiveDate": null,
        "listPrice": 42600,
        "netPrice": 23500,
        "costPrice": 18240
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 37000,
        "netPrice": 21000,
        "costPrice": 15840
      }
    ]
  },
  {
    "id": "REN-07-035",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "LBS用PF 7.2kV T88A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 63100,
    "netPrice": 34705,
    "costPrice": 27060,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "三菱15%up",
        "effectiveDate": null,
        "listPrice": 48500,
        "netPrice": 26700,
        "costPrice": 20820
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 42200,
        "netPrice": 24000,
        "costPrice": 18120
      }
    ]
  },
  {
    "id": "REN-07-036",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "LBS用PF 7.2kV G150A (T100A)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 70700,
    "netPrice": 38885,
    "costPrice": 18600,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "三菱15%up",
        "effectiveDate": null,
        "listPrice": 54400,
        "netPrice": 30000,
        "costPrice": 14320
      }
    ]
  },
  {
    "id": "REN-07-039",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "ｴﾈｾｰﾊﾞ 7.2kV 200A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1037400,
    "netPrice": 570570,
    "costPrice": 520000,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-07-042",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "OPTr 6600/110V 600VA",
    "spec": "EMT-BB",
    "defaultQty": 1,
    "listPrice": 314000,
    "netPrice": 173000,
    "costPrice": 107400,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-07-045",
    "categoryId": "cat-6",
    "category": "LBS",
    "name": "ｴﾅﾐｯｸ 7.2kV 200A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 775600,
    "netPrice": 460000,
    "costPrice": 460000,
    "priceLabel": "30%up R8.7.1(確定価格)",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "三菱15%up",
        "effectiveDate": null,
        "listPrice": 798000,
        "netPrice": 400000,
        "costPrice": 400000,
        "sourceName": "ｴﾈｾｰﾊﾞ 7.2kV 200A"
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 798000,
        "netPrice": 400000,
        "costPrice": 400000,
        "sourceName": "ｴﾈｾｰﾊﾞ 7.2kV 200A"
      }
    ]
  },
  {
    "id": "REN-08-003",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC 7.2kV 30A",
    "spec": "宇野",
    "defaultQty": 1,
    "listPrice": 21700,
    "netPrice": 14400,
    "costPrice": 14400,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-004",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC 7.2kV 30A (SC)",
    "spec": "宇野",
    "defaultQty": 1,
    "listPrice": 18200,
    "netPrice": 12200,
    "costPrice": 12200,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-005",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC 7.2kV 50A",
    "spec": "宇野",
    "defaultQty": 1,
    "listPrice": 28400,
    "netPrice": 19000,
    "costPrice": 19000,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-006",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC 7.2kV 30A",
    "spec": "宇野",
    "defaultQty": 1,
    "listPrice": 25600,
    "netPrice": 17000,
    "costPrice": 17000,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-007",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC 7.2kV 50A",
    "spec": "宇野",
    "defaultQty": 1,
    "listPrice": 33100,
    "netPrice": 21800,
    "costPrice": 21800,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-009",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用ﾃﾝｼｮﾝF 7.2kV 5A",
    "spec": "宇野",
    "defaultQty": 1,
    "listPrice": 1500,
    "netPrice": 2000,
    "costPrice": 1200,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-010",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用ﾃﾝｼｮﾝF 7.2kV 10A",
    "spec": "宇野",
    "defaultQty": 1,
    "listPrice": 1500,
    "netPrice": 2000,
    "costPrice": 1200,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-011",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用ﾃﾝｼｮﾝF 7.2kV 15A",
    "spec": "宇野",
    "defaultQty": 1,
    "listPrice": 2200,
    "netPrice": 2000,
    "costPrice": 1600,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-012",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用ﾃﾝｼｮﾝF 7.2kV 20A",
    "spec": "宇野",
    "defaultQty": 1,
    "listPrice": 2200,
    "netPrice": 2000,
    "costPrice": 1600,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-013",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用ﾃﾝｼｮﾝF 7.2kV 30A",
    "spec": "宇野",
    "defaultQty": 1,
    "listPrice": 2200,
    "netPrice": 2000,
    "costPrice": 1600,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-014",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用ﾃﾝｼｮﾝF 7.2kV 50A",
    "spec": "宇野",
    "defaultQty": 1,
    "listPrice": 4500,
    "netPrice": 3000,
    "costPrice": 3000,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-015",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用ﾀｲﾑﾗｸﾞｸﾞF 7.2kV 5A",
    "spec": "宇野",
    "defaultQty": 1,
    "listPrice": 3100,
    "netPrice": 2000,
    "costPrice": 1800,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-016",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用ﾀｲﾑﾗｸﾞｸﾞF 7.2kV 10A",
    "spec": "宇野",
    "defaultQty": 1,
    "listPrice": 3100,
    "netPrice": 2000,
    "costPrice": 1800,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-017",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用ﾀｲﾑﾗｸﾞｸﾞF 7.2kV 15A",
    "spec": "宇野",
    "defaultQty": 1,
    "listPrice": 4300,
    "netPrice": 3000,
    "costPrice": 2900,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-018",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用ﾀｲﾑﾗｸﾞｸﾞF 7.2kV 20A",
    "spec": "宇野",
    "defaultQty": 1,
    "listPrice": 4300,
    "netPrice": 3000,
    "costPrice": 2900,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-019",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用ﾀｲﾑﾗｸﾞｸﾞF 7.2kV 30A",
    "spec": "宇野",
    "defaultQty": 1,
    "listPrice": 5000,
    "netPrice": 4000,
    "costPrice": 3700,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-021",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用PF 7.2kV G10A",
    "spec": "宇野",
    "defaultQty": 1,
    "listPrice": 19000,
    "netPrice": 14000,
    "costPrice": 13200,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-022",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用PF 7.2kV G15A",
    "spec": "宇野",
    "defaultQty": 1,
    "listPrice": 19800,
    "netPrice": 16000,
    "costPrice": 15300,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-024",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "LA 断路形避雷器 8.4KV 2.5KA",
    "spec": "宇野",
    "defaultQty": 1,
    "listPrice": 111000,
    "netPrice": 66000,
    "costPrice": 66000,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-028",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "LA 8.4kV 2.5kA 単極",
    "spec": "宇野",
    "defaultQty": 1,
    "listPrice": 19000,
    "netPrice": 12000,
    "costPrice": 11900,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-032",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC 7.2kV 30A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 22000,
    "netPrice": 12000,
    "costPrice": 11700,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-033",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC 7.2kV 30A (SC)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 17400,
    "netPrice": 11000,
    "costPrice": 9600,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-034",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC 7.2kV 50A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 28500,
    "netPrice": 16000,
    "costPrice": 15300,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-035",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC 7.2kV 100A",
    "spec": "宇野×南〇",
    "defaultQty": 1,
    "listPrice": 46400,
    "netPrice": 25000,
    "costPrice": 24900,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-036",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC 7.2kV 30A",
    "spec": "南",
    "defaultQty": 1,
    "listPrice": 23000,
    "netPrice": 13000,
    "costPrice": 12300,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-037",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC 7.2kV 30A(SC)",
    "spec": "南",
    "defaultQty": 1,
    "listPrice": 19100,
    "netPrice": 9000,
    "costPrice": 8800,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-038",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC 7.2kV 50A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 26300,
    "netPrice": 14000,
    "costPrice": 13900,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-040",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用ﾃﾝｼｮﾝF 7.2kV 5A",
    "spec": "南",
    "defaultQty": 1,
    "listPrice": 1720,
    "netPrice": 1000,
    "costPrice": 940,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-041",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用ﾃﾝｼｮﾝF 7.2kV 10A",
    "spec": "南",
    "defaultQty": 1,
    "listPrice": 1720,
    "netPrice": 1000,
    "costPrice": 940,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-042",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用ﾃﾝｼｮﾝF 7.2kV 15A",
    "spec": "南",
    "defaultQty": 1,
    "listPrice": 2440,
    "netPrice": 2000,
    "costPrice": 1300,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-043",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用ﾃﾝｼｮﾝF 7.2kV 20A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2030,
    "netPrice": 2000,
    "costPrice": 1100,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-044",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用ﾃﾝｼｮﾝF 7.2kV 30A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2030,
    "netPrice": 2000,
    "costPrice": 1100,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-045",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用ﾃﾝｼｮﾝF 7.2kV 50A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2760,
    "netPrice": 2000,
    "costPrice": 1400,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-047",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用ﾃﾝｼｮﾝF 7.2kV 50A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3040,
    "netPrice": 2000,
    "costPrice": 1600,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-048",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用ﾃﾝｼｮﾝF 7.2kV 75A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3520,
    "netPrice": 2000,
    "costPrice": 1900,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-049",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用ﾃﾝｼｮﾝF 7.2kV 100A",
    "spec": "宇野",
    "defaultQty": 1,
    "listPrice": 4320,
    "netPrice": 3000,
    "costPrice": 2800,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-051",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用ﾀｲﾑﾗｸﾞF 7.2kV 5A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1990,
    "netPrice": 1000,
    "costPrice": 1060,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-052",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用ﾀｲﾑﾗｸﾞF 7.2kV 10A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1990,
    "netPrice": 1000,
    "costPrice": 1060,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-053",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用ﾀｲﾑﾗｸﾞF 7.2kV 15A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2730,
    "netPrice": 2000,
    "costPrice": 1450,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-054",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用ﾀｲﾑﾗｸﾞF 7.2kV 20A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2730,
    "netPrice": 2000,
    "costPrice": 1450,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-055",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用ﾀｲﾑﾗｸﾞF 7.2kV 30A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3380,
    "netPrice": 2000,
    "costPrice": 1900,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-057",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用PF 7.2kV G7A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 17000,
    "netPrice": 11000,
    "costPrice": 10100,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-058",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用PF 7.2kV G10A",
    "spec": "南",
    "defaultQty": 1,
    "listPrice": 22400,
    "netPrice": 13000,
    "costPrice": 12700,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-059",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "PC用PF 7.2kV G15A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 21200,
    "netPrice": 13000,
    "costPrice": 12100,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-061",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "直結型PF 7.2kV 15A",
    "spec": "南",
    "defaultQty": 1,
    "listPrice": 25100,
    "netPrice": 18000,
    "costPrice": 17400,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-062",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "直結型PF 7.2kV 25A",
    "spec": "南",
    "defaultQty": 1,
    "listPrice": 29900,
    "netPrice": 21000,
    "costPrice": 20400,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-063",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "直結型PF 7.2kV 40A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 52100,
    "netPrice": 34000,
    "costPrice": 33200,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-08-066",
    "categoryId": "cat-7",
    "category": "PC・LA",
    "name": "素通し線 (30A)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1620,
    "netPrice": 1000,
    "costPrice": 890,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-09-005",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 0.766kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2277000,
    "netPrice": 223000,
    "costPrice": 222550,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 1981000,
        "netPrice": 195000,
        "costPrice": 193600
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 1722000,
        "netPrice": 169000,
        "costPrice": 168300,
        "sourceName": "SR 6.6kV 0.766kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-006",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 1.15kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2277000,
    "netPrice": 223000,
    "costPrice": 222550,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 1981000,
        "netPrice": 195000,
        "costPrice": 193600
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 1722000,
        "netPrice": 169000,
        "costPrice": 168300,
        "sourceName": "SR 6.6kV 1.15kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-007",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 1.53kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2277000,
    "netPrice": 223000,
    "costPrice": 222550,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 1981000,
        "netPrice": 195000,
        "costPrice": 193600
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 1722000,
        "netPrice": 169000,
        "costPrice": 168300,
        "sourceName": "SR 6.6kV 1.53kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-008",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 1.91kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2277000,
    "netPrice": 223000,
    "costPrice": 222550,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 1981000,
        "netPrice": 195000,
        "costPrice": 193600
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 1722000,
        "netPrice": 169000,
        "costPrice": 168300,
        "sourceName": "SR 6.6kV 1.91kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-009",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 2.30kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2485000,
    "netPrice": 244000,
    "costPrice": 243010,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 2161000,
        "netPrice": 212000,
        "costPrice": 211400
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 1879000,
        "netPrice": 184000,
        "costPrice": 183740,
        "sourceName": "SR 6.6kV 2.30kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-010",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 3.19kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2705000,
    "netPrice": 265000,
    "costPrice": 264400,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 2352000,
        "netPrice": 230000,
        "costPrice": 230000
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 2045000,
        "netPrice": 200000,
        "costPrice": 199980,
        "sourceName": "SR 6.6kV 3.19kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-011",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 4.79kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2875000,
    "netPrice": 282000,
    "costPrice": 281160,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 2501000,
        "netPrice": 245000,
        "costPrice": 244400
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 2174000,
        "netPrice": 213000,
        "costPrice": 212520,
        "sourceName": "SR 6.6kV 4.79kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-012",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 6.38kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3075000,
    "netPrice": 301000,
    "costPrice": 300700,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 2674000,
        "netPrice": 263000,
        "costPrice": 261600
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 2325000,
        "netPrice": 228000,
        "costPrice": 227440,
        "sourceName": "SR 6.6kV 6.38kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-013",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 9.57kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3416000,
    "netPrice": 335000,
    "costPrice": 334090,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 2971000,
        "netPrice": 291000,
        "costPrice": 290600
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 2583000,
        "netPrice": 253000,
        "costPrice": 252650,
        "sourceName": "SR 6.6kV 9.57kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-014",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 12.8kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3773000,
    "netPrice": 370000,
    "costPrice": 369070,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 3281000,
        "netPrice": 322000,
        "costPrice": 321000
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 2853000,
        "netPrice": 280000,
        "costPrice": 279050,
        "sourceName": "SR 6.6kV 12.8kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-015",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 16.0kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 4247000,
    "netPrice": 416000,
    "costPrice": 415140,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 3693000,
        "netPrice": 362000,
        "costPrice": 361000
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 3211000,
        "netPrice": 314000,
        "costPrice": 313900,
        "sourceName": "SR 6.6kV 16.0kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-016",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 19.1kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 4613000,
    "netPrice": 452000,
    "costPrice": 451180,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 4012000,
        "netPrice": 394000,
        "costPrice": 392300
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 3488000,
        "netPrice": 342000,
        "costPrice": 341090,
        "sourceName": "SR 6.6kV 19.1kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-017",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 25.5kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 5854000,
    "netPrice": 573000,
    "costPrice": 572480,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 5090000,
        "netPrice": 498000,
        "costPrice": 497800
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 4426000,
        "netPrice": 433000,
        "costPrice": 432830,
        "sourceName": "SR 6.6kV 25.5kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-018",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 31.9kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 6936000,
    "netPrice": 679000,
    "costPrice": 678220,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 6031000,
        "netPrice": 590000,
        "costPrice": 589800
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 5244000,
        "netPrice": 513000,
        "costPrice": 512820,
        "sourceName": "SR 6.6kV 31.9kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-021",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SC 7020V 12.8kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 286000,
    "netPrice": 28000,
    "costPrice": 27850,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 268000,
        "netPrice": 27000,
        "costPrice": 26000
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 243000,
        "netPrice": 24000,
        "costPrice": 23630,
        "sourceName": "SC 6.6kV 12.8kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-022",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SC 7020V 19.1kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 322000,
    "netPrice": 32000,
    "costPrice": 31420,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 302000,
        "netPrice": 30000,
        "costPrice": 29400
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 274000,
        "netPrice": 27000,
        "costPrice": 26660,
        "sourceName": "SC 6.6kV 19.1kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-023",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SC 7020V 25.5kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 367000,
    "netPrice": 36000,
    "costPrice": 35900,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 344000,
        "netPrice": 35000,
        "costPrice": 33600
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 312000,
        "netPrice": 31000,
        "costPrice": 30490,
        "sourceName": "SC 6.6kV 25.5kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-024",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SC 7020V 31.9kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 403000,
    "netPrice": 40000,
    "costPrice": 39470,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 378000,
        "netPrice": 38000,
        "costPrice": 36900
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 343000,
        "netPrice": 34000,
        "costPrice": 33530,
        "sourceName": "SC 6.6kV 31.9kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-025",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SC 7020V 38.3kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 437000,
    "netPrice": 43000,
    "costPrice": 42770,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 409000,
        "netPrice": 41000,
        "costPrice": 40000
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 371000,
        "netPrice": 37000,
        "costPrice": 36300,
        "sourceName": "SC 6.6kV 38.3kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-026",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SC 7020V 53.2kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 541000,
    "netPrice": 53000,
    "costPrice": 52800,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 506000,
        "netPrice": 50000,
        "costPrice": 49400
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 460000,
        "netPrice": 45000,
        "costPrice": 44880,
        "sourceName": "SC 6.6kV 53.2kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-027",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SC 7020V 79.8kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 677000,
    "netPrice": 67000,
    "costPrice": 66260,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 633000,
        "netPrice": 63000,
        "costPrice": 61900
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 575000,
        "netPrice": 57000,
        "costPrice": 56230,
        "sourceName": "SC 6.6kV 79.8kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-028",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SC 7020V 106kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 777000,
    "netPrice": 77000,
    "costPrice": 76030,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 726000,
        "netPrice": 72000,
        "costPrice": 71100
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 660000,
        "netPrice": 65000,
        "costPrice": 64550,
        "sourceName": "SC 6.6kV 106kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-029",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SC 7020V 160kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1058000,
    "netPrice": 104000,
    "costPrice": 103490,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 989000,
        "netPrice": 97000,
        "costPrice": 96800
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 899000,
        "netPrice": 88000,
        "costPrice": 87910,
        "sourceName": "SC 6.6kV 160kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-030",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SC 7020V 213kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1330000,
    "netPrice": 131000,
    "costPrice": 130020,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 1243000,
        "netPrice": 123000,
        "costPrice": 121600
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 1130000,
        "netPrice": 111000,
        "costPrice": 110480,
        "sourceName": "SC 6.6kV 213kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-031",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SC 7020V 266kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1663000,
    "netPrice": 163000,
    "costPrice": 162760,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 1555000,
        "netPrice": 153000,
        "costPrice": 152100
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 1413000,
        "netPrice": 139000,
        "costPrice": 138200,
        "sourceName": "SC 6.6kV 266kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-032",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SC 7020V 319kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1950000,
    "netPrice": 191000,
    "costPrice": 190740,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 1822000,
        "netPrice": 179000,
        "costPrice": 178200
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 1656000,
        "netPrice": 162000,
        "costPrice": 161960,
        "sourceName": "SC 6.6kV 319kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-033",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SC 7020V 426kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2600000,
    "netPrice": 255000,
    "costPrice": 254230,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 2430000,
        "netPrice": 238000,
        "costPrice": 237600
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 2209000,
        "netPrice": 216000,
        "costPrice": 215950,
        "sourceName": "SC 6.6kV 426kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-034",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SC 7020V 532kvar 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3243000,
    "netPrice": 318000,
    "costPrice": 317060,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 3031000,
        "netPrice": 297000,
        "costPrice": 296400
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 2755000,
        "netPrice": 270000,
        "costPrice": 269410,
        "sourceName": "SC 6.6kV 532kvar 60Hz 油入形"
      }
    ]
  },
  {
    "id": "REN-09-036",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "保護検出器",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 9500,
    "netPrice": 6000,
    "costPrice": 5940,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 10000,
        "netPrice": 7000,
        "costPrice": 6000
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 8300,
        "netPrice": 6000,
        "costPrice": 5410
      }
    ]
  },
  {
    "id": "REN-09-037",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "保護検出器",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 9500,
    "netPrice": 9000,
    "costPrice": 8470,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 10000,
        "netPrice": 9000,
        "costPrice": 8500
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 8300,
        "netPrice": 8000,
        "costPrice": 7710
      }
    ]
  },
  {
    "id": "REN-09-044",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 0.766kvar 60Hz 乾式ﾓｰﾙﾄﾞ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2792000,
    "netPrice": 274000,
    "costPrice": 273110,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 2428000,
        "netPrice": 235000,
        "costPrice": 234100,
        "sourceName": "SR 6600V 1.91kvar 60Hz 乾式モールド"
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 2111000,
        "netPrice": 204000,
        "costPrice": 203500,
        "sourceName": "SR 6.6kV 1.91kvar 60Hz 乾式モールド"
      }
    ]
  },
  {
    "id": "REN-09-045",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 1.15kvar 60Hz 乾式ﾓｰﾙﾄﾞ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2792000,
    "netPrice": 274000,
    "costPrice": 273110,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-09-046",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 1.53kvar 60Hz 乾式ﾓｰﾙﾄﾞ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2792000,
    "netPrice": 274000,
    "costPrice": 273110,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-09-047",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 1.91kvar 60Hz 乾式ﾓｰﾙﾄﾞ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2792000,
    "netPrice": 274000,
    "costPrice": 273110,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 2938000,
        "netPrice": 295000,
        "costPrice": 293800,
        "sourceName": "SR 6600V 3.19kvar 60Hz 乾式モールド"
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 2554000,
        "netPrice": 256000,
        "costPrice": 255400,
        "sourceName": "SR 6.6kV 3.19kvar 60Hz 乾式モールド"
      }
    ]
  },
  {
    "id": "REN-09-048",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 2.30kvar 60Hz 乾式ﾓｰﾙﾄﾞ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3191000,
    "netPrice": 312000,
    "costPrice": 311920,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 3119000,
        "netPrice": 356000,
        "costPrice": 355740,
        "sourceName": "SR 6600V 4.79kvar 60Hz 乾式モールド"
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 2712000,
        "netPrice": 266000,
        "costPrice": 265190,
        "sourceName": "SR 6.6kV 4.79kvar 60Hz 乾式モールド"
      }
    ]
  },
  {
    "id": "REN-09-049",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 3.19kvar 60Hz 乾式ﾓｰﾙﾄﾞ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3378000,
    "netPrice": 331000,
    "costPrice": 330260,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 3312000,
        "netPrice": 332000,
        "costPrice": 331200,
        "sourceName": "SR 6600V 6.38kvar 60Hz 乾式モールド"
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 2880000,
        "netPrice": 288000,
        "costPrice": 288000,
        "sourceName": "SR 6.6kV 6.38kvar 60Hz 乾式モールド"
      }
    ]
  },
  {
    "id": "REN-09-050",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 4.79kvar 60Hz 乾式ﾓｰﾙﾄﾞ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3587000,
    "netPrice": 351000,
    "costPrice": 350720,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-09-051",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 6.38kvar 60Hz 乾式ﾓｰﾙﾄﾞ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3809000,
    "netPrice": 373000,
    "costPrice": 372370,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-09-052",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 9.57kvar 60Hz 乾式ﾓｰﾙﾄﾞ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 4238000,
    "netPrice": 415000,
    "costPrice": 414480,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-09-053",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 12.8kvar 60Hz 乾式ﾓｰﾙﾄﾞ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 5114000,
    "netPrice": 501000,
    "costPrice": 500150,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "SR15％ SC10％ 2024.10",
        "effectiveDate": "2024-10-01",
        "listPrice": 616000,
        "netPrice": 275000,
        "costPrice": 275000,
        "sourceName": "LB5-S形 LB522B6018S26"
      }
    ]
  },
  {
    "id": "REN-09-054",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 16.0kvar 60Hz 乾式ﾓｰﾙﾄﾞ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 5981000,
    "netPrice": 585000,
    "costPrice": 584890,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-09-055",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 19.1kvar 60Hz 乾式ﾓｰﾙﾄﾞ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 6844000,
    "netPrice": 670000,
    "costPrice": 669110,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 308000,
        "netPrice": 92000,
        "costPrice": 92000,
        "sourceName": "低圧ｺﾝﾃﾞﾝｻ 220/60Hz 25.5kvar 油入式 L=6%対応"
      }
    ]
  },
  {
    "id": "REN-09-056",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 25.5kvar 60Hz 乾式ﾓｰﾙﾄﾞ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 8394000,
    "netPrice": 821000,
    "costPrice": 820640,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-09-057",
    "categoryId": "cat-8",
    "category": "SC・SR",
    "name": "SR 6600V 31.9kvar 60Hz 乾式ﾓｰﾙﾄﾞ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 9781000,
    "netPrice": 957000,
    "costPrice": 956470,
    "priceLabel": "SR15％ SC７％ 2026.7",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-10-004",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 10kVA 6600/210-105V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1507000,
    "netPrice": 449600,
    "costPrice": 449600,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 792000,
        "netPrice": null,
        "costPrice": 158100,
        "sourceName": "3OT1P10K66K2"
      }
    ]
  },
  {
    "id": "REN-10-005",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 20kVA 6600/210-105V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1765000,
    "netPrice": 526800,
    "costPrice": 526800,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 928000,
        "netPrice": null,
        "costPrice": 207000,
        "sourceName": "3OT1P20K66K2"
      }
    ]
  },
  {
    "id": "REN-10-006",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 30kVA 6600/210-105V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2004000,
    "netPrice": 598000,
    "costPrice": 598000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1054000,
        "netPrice": null,
        "costPrice": 229500,
        "sourceName": "3OT1P30K66K2"
      }
    ]
  },
  {
    "id": "REN-10-007",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 50kVA 6600/210-105V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2458000,
    "netPrice": 733600,
    "costPrice": 733600,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1294000,
        "netPrice": null,
        "costPrice": 287700,
        "sourceName": "3OT1P50K66K2"
      }
    ]
  },
  {
    "id": "REN-10-008",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 75kVA 6600/210-105V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3045000,
    "netPrice": 908800,
    "costPrice": 908800,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1600000,
        "netPrice": null,
        "costPrice": 347400,
        "sourceName": "3OT1P75K66K2"
      }
    ]
  },
  {
    "id": "REN-10-009",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 100kVA 6600/210-105V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3675000,
    "netPrice": 1096800,
    "costPrice": 1096800,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1933000,
        "netPrice": null,
        "costPrice": 409500,
        "sourceName": "3OT1P100K66K2"
      }
    ]
  },
  {
    "id": "REN-10-010",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 150kVA 6600/210-105V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 4916000,
    "netPrice": 1467200,
    "costPrice": 1467200,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 2587000,
        "netPrice": null,
        "costPrice": 536100,
        "sourceName": "3OT1P150K66K2"
      }
    ]
  },
  {
    "id": "REN-10-011",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 200kVA 6600/210-105V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 6080000,
    "netPrice": 1814800,
    "costPrice": 1814800,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 3199000,
        "netPrice": null,
        "costPrice": 687000,
        "sourceName": "3OT1P200K66K2"
      }
    ]
  },
  {
    "id": "REN-10-012",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 300kVA 6600/210-105V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 8469000,
    "netPrice": 2528000,
    "costPrice": 2528000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 4457000,
        "netPrice": null,
        "costPrice": 936600,
        "sourceName": "3OT1P300K66K2"
      }
    ]
  },
  {
    "id": "REN-10-013",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 500kVA 6600/210-105V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 13890000,
    "netPrice": 4146000,
    "costPrice": 4146000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 7310000,
        "netPrice": null,
        "costPrice": 1499100,
        "sourceName": "3OT1P500K66K2"
      }
    ]
  },
  {
    "id": "REN-10-016",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 10kVA 6600/210-105V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1655000,
    "netPrice": 494000,
    "costPrice": 494000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 871000,
        "netPrice": null,
        "costPrice": 211300
      }
    ]
  },
  {
    "id": "REN-10-017",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 20kVA 6600/210-105V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1942000,
    "netPrice": 579600,
    "costPrice": 579600,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1022000,
        "netPrice": null,
        "costPrice": 247600
      }
    ]
  },
  {
    "id": "REN-10-018",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 30kVA 6600/210-105V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2203000,
    "netPrice": 657600,
    "costPrice": 657600,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1160000,
        "netPrice": null,
        "costPrice": 281000
      }
    ]
  },
  {
    "id": "REN-10-019",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 50kVA 6600/210-105V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2705000,
    "netPrice": 807200,
    "costPrice": 807200,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1422000,
        "netPrice": null,
        "costPrice": 344800
      }
    ]
  },
  {
    "id": "REN-10-020",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 75kVA 6600/210-105V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3349000,
    "netPrice": 999600,
    "costPrice": 999600,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1763000,
        "netPrice": null,
        "costPrice": 427300
      }
    ]
  },
  {
    "id": "REN-10-021",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 100kVA 6600/210-105V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 4046000,
    "netPrice": 1207600,
    "costPrice": 1207600,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 2128000,
        "netPrice": null,
        "costPrice": 516000
      }
    ]
  },
  {
    "id": "REN-10-022",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 150kVA 6600/210-105V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 5406000,
    "netPrice": 1613600,
    "costPrice": 1613600,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 2844000,
        "netPrice": null,
        "costPrice": 689600
      }
    ]
  },
  {
    "id": "REN-10-023",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 200kVA 6600/210-105V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 6687000,
    "netPrice": 1996000,
    "costPrice": 1996000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 3519000,
        "netPrice": null,
        "costPrice": 853400
      }
    ]
  },
  {
    "id": "REN-10-024",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 300kVA 6600/210-105V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 9320000,
    "netPrice": 2782000,
    "costPrice": 2782000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 4904000,
        "netPrice": null,
        "costPrice": 1189000
      }
    ]
  },
  {
    "id": "REN-10-025",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 500kVA 6600/210-105V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 15281000,
    "netPrice": 4561200,
    "costPrice": 4561200,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 8042000,
        "netPrice": null,
        "costPrice": 1950000
      }
    ]
  },
  {
    "id": "REN-10-028",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 20kVA 6600/210V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2178000,
    "netPrice": 650000,
    "costPrice": 650000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1146000,
        "netPrice": null,
        "costPrice": 245700,
        "sourceName": "3OT3P20K66K2"
      }
    ]
  },
  {
    "id": "REN-10-029",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 30kVA 6600/210V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2395000,
    "netPrice": 714800,
    "costPrice": 714800,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1260000,
        "netPrice": null,
        "costPrice": 284400,
        "sourceName": "3OT3P30K66K2"
      }
    ]
  },
  {
    "id": "REN-10-030",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 50kVA 6600/210V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2875000,
    "netPrice": 858000,
    "costPrice": 858000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1512000,
        "netPrice": null,
        "costPrice": 338100,
        "sourceName": "3OT3P50K66K2"
      }
    ]
  },
  {
    "id": "REN-10-031",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 75kVA 6600/210V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3438000,
    "netPrice": 1026000,
    "costPrice": 1026000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1808000,
        "netPrice": null,
        "costPrice": 394200,
        "sourceName": "3OT3P75K66K2"
      }
    ]
  },
  {
    "id": "REN-10-032",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 100kVA 6600/210V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 4046000,
    "netPrice": 1207600,
    "costPrice": 1207600,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 2128000,
        "netPrice": null,
        "costPrice": 455100,
        "sourceName": "3OT3P100K66K2"
      }
    ]
  },
  {
    "id": "REN-10-033",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 150kVA 6600/210V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 5216000,
    "netPrice": 1556800,
    "costPrice": 1556800,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 2743000,
        "netPrice": null,
        "costPrice": 638400,
        "sourceName": "3OT3P150K66K2"
      }
    ]
  },
  {
    "id": "REN-10-034",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 200kVA 6600/210V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 6344000,
    "netPrice": 1893600,
    "costPrice": 1893600,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 3338000,
        "netPrice": null,
        "costPrice": 746700,
        "sourceName": "3OT3P200K66K2"
      }
    ]
  },
  {
    "id": "REN-10-035",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 300kVA 6600/210V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 9324000,
    "netPrice": 2783200,
    "costPrice": 2783200,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 4906000,
        "netPrice": null,
        "costPrice": 1047900,
        "sourceName": "3OT3P300K66K2"
      }
    ]
  },
  {
    "id": "REN-10-036",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 500kVA 6600/210V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 14852000,
    "netPrice": 4433200,
    "costPrice": 4433200,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 7815000,
        "netPrice": null,
        "costPrice": 1633200,
        "sourceName": "3OT3P500K66K2"
      }
    ]
  },
  {
    "id": "REN-10-037",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 750kVA 6600/210V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 19045000,
    "netPrice": 5684800,
    "costPrice": 5684800,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 11541000,
        "netPrice": null,
        "costPrice": 2406900,
        "sourceName": "3OT3P750K66K2"
      }
    ]
  },
  {
    "id": "REN-10-038",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 1000kVA 6600/210V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 25389000,
    "netPrice": 7578800,
    "costPrice": 7578800,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 15386000,
        "netPrice": null,
        "costPrice": 3090000,
        "sourceName": "3OT3P1000K66K2"
      }
    ]
  },
  {
    "id": "REN-10-039",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 1500kVA 6600/210V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 38284000,
    "netPrice": 11428000,
    "costPrice": 11428000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 23201000,
        "netPrice": null,
        "costPrice": null,
        "sourceName": "3OT3P1500K66K2"
      }
    ]
  },
  {
    "id": "REN-10-040",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 2000kVA 6600/210V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 50777000,
    "netPrice": 15157200,
    "costPrice": 15157200,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 30772000,
        "netPrice": null,
        "costPrice": null,
        "sourceName": "3OT3P2000K66K2"
      }
    ]
  },
  {
    "id": "REN-10-043",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 20kVA 6600/210V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2395000,
    "netPrice": 714800,
    "costPrice": 714800,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1260000,
        "netPrice": null,
        "costPrice": 305500
      }
    ]
  },
  {
    "id": "REN-10-044",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 30kVA 6600/210V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2638000,
    "netPrice": 787200,
    "costPrice": 787200,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1387000,
        "netPrice": null,
        "costPrice": 336300
      }
    ]
  },
  {
    "id": "REN-10-045",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 50kVA 6600/210V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3159000,
    "netPrice": 942800,
    "costPrice": 942800,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1661000,
        "netPrice": null,
        "costPrice": 402600
      }
    ]
  },
  {
    "id": "REN-10-046",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 75kVA 6600/210V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3786000,
    "netPrice": 1130000,
    "costPrice": 1130000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1992000,
        "netPrice": null,
        "costPrice": 482900
      }
    ]
  },
  {
    "id": "REN-10-047",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 100kVA 6600/210V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 4448000,
    "netPrice": 1327600,
    "costPrice": 1327600,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 2340000,
        "netPrice": null,
        "costPrice": 567400
      }
    ]
  },
  {
    "id": "REN-10-048",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 150kVA 6600/210V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 5736000,
    "netPrice": 1712000,
    "costPrice": 1712000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 3018000,
        "netPrice": null,
        "costPrice": 731900
      }
    ]
  },
  {
    "id": "REN-10-049",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 200kVA 6600/210V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 6974000,
    "netPrice": 2081600,
    "costPrice": 2081600,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 3669000,
        "netPrice": null,
        "costPrice": 889800
      }
    ]
  },
  {
    "id": "REN-10-050",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 300kVA 6600/210V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 10258000,
    "netPrice": 3062000,
    "costPrice": 3062000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 5398000,
        "netPrice": null,
        "costPrice": 1309000
      }
    ]
  },
  {
    "id": "REN-10-051",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 500kVA 6600/210V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 16338000,
    "netPrice": 4876800,
    "costPrice": 4876800,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 8598000,
        "netPrice": null,
        "costPrice": 2085000
      }
    ]
  },
  {
    "id": "REN-10-052",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 750kVA 6600/210V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 20950000,
    "netPrice": 6253600,
    "costPrice": 6253600,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 12696000,
        "netPrice": null,
        "costPrice": 3079000
      }
    ]
  },
  {
    "id": "REN-10-053",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 1000kVA 6600/210V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 27930000,
    "netPrice": 8337200,
    "costPrice": 8337200,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 16926000,
        "netPrice": null,
        "costPrice": 4105000
      }
    ]
  },
  {
    "id": "REN-10-054",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 1500kVA 6600/210V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-055",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 2000kVA 6600/210V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-058",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 20kVA 6600/440-254V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1374000,
        "netPrice": null,
        "costPrice": 282000
      }
    ]
  },
  {
    "id": "REN-10-059",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 30kVA 6600/440-254V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1513000,
        "netPrice": null,
        "costPrice": 326700
      }
    ]
  },
  {
    "id": "REN-10-060",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 50kVA 6600/440-254V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1812000,
        "netPrice": null,
        "costPrice": 389100
      }
    ]
  },
  {
    "id": "REN-10-061",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 75kVA 6600/440-254V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3951000,
    "netPrice": 1179200,
    "costPrice": 1179200,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 2079000,
        "netPrice": null,
        "costPrice": 453000
      }
    ]
  },
  {
    "id": "REN-10-062",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 100kVA 6600/440-254V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 4650000,
    "netPrice": 1388000,
    "costPrice": 1388000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 2447000,
        "netPrice": null,
        "costPrice": 523500
      }
    ]
  },
  {
    "id": "REN-10-063",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 150kVA 6600/440-254V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 5994000,
    "netPrice": 1789200,
    "costPrice": 1789200,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 3154000,
        "netPrice": null,
        "costPrice": 734100
      }
    ]
  },
  {
    "id": "REN-10-064",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 200kVA 6600/440-254V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 7299000,
    "netPrice": 2178800,
    "costPrice": 2178800,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 3841000,
        "netPrice": null,
        "costPrice": 858000
      }
    ]
  },
  {
    "id": "REN-10-065",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 300kVA 6600/440-254V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 10258000,
    "netPrice": 3062000,
    "costPrice": 3062000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 5398000,
        "netPrice": null,
        "costPrice": 1205400
      }
    ]
  },
  {
    "id": "REN-10-066",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 500kVA 6600/440-254V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 16338000,
    "netPrice": 4876800,
    "costPrice": 4876800,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 8598000,
        "netPrice": null,
        "costPrice": 1878000
      }
    ]
  },
  {
    "id": "REN-10-067",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 750kVA 6600/440-254V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 19997000,
    "netPrice": 5969200,
    "costPrice": 5969200,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 12118000,
        "netPrice": null,
        "costPrice": 2647500
      }
    ]
  },
  {
    "id": "REN-10-068",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 1000kVA 6600/440-254V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 26660000,
    "netPrice": 7958000,
    "costPrice": 7958000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 16157000,
        "netPrice": null,
        "costPrice": 3399000
      }
    ]
  },
  {
    "id": "REN-10-069",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 1500kVA 6600/440-254V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 36372000,
    "netPrice": 10857200,
    "costPrice": 10857200,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 22042000,
        "netPrice": null,
        "costPrice": null
      }
    ]
  },
  {
    "id": "REN-10-070",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 2000kVA 6600/440-254V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 48238000,
    "netPrice": 14399200,
    "costPrice": 14399200,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 29234000,
        "netPrice": null,
        "costPrice": null
      }
    ]
  },
  {
    "id": "REN-10-073",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 20kVA 6600/440-254V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1512000,
        "netPrice": null,
        "costPrice": 310500
      }
    ]
  },
  {
    "id": "REN-10-074",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 30kVA 6600/440-254V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1663000,
        "netPrice": null,
        "costPrice": 359700
      }
    ]
  },
  {
    "id": "REN-10-075",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 50kVA 6600/440-254V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1996000,
        "netPrice": null,
        "costPrice": 427500
      }
    ]
  },
  {
    "id": "REN-10-076",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 75kVA 6600/440-254V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 2288000,
        "netPrice": null,
        "costPrice": 498900
      }
    ]
  },
  {
    "id": "REN-10-077",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 100kVA 6600/440-254V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 2693000,
        "netPrice": null,
        "costPrice": 575700
      }
    ]
  },
  {
    "id": "REN-10-078",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 150kVA 6600/440-254V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 3471000,
        "netPrice": null,
        "costPrice": 807600
      }
    ]
  },
  {
    "id": "REN-10-079",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 200kVA 6600/440-254V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 4224000,
        "netPrice": null,
        "costPrice": 944400
      }
    ]
  },
  {
    "id": "REN-10-080",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 300kVA 6600/440-254V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 11283000,
    "netPrice": 3368000,
    "costPrice": 3368000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 5938000,
        "netPrice": null,
        "costPrice": 1325700
      }
    ]
  },
  {
    "id": "REN-10-081",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 500kVA 6600/440-254V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 9457000,
        "netPrice": null,
        "costPrice": 2065500
      }
    ]
  },
  {
    "id": "REN-10-082",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 750kVA 6600/440-254V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 13331000,
        "netPrice": null,
        "costPrice": 2912400
      }
    ]
  },
  {
    "id": "REN-10-083",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 1000kVA 6600/440-254V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 17773000,
        "netPrice": null,
        "costPrice": 3738000
      }
    ]
  },
  {
    "id": "REN-10-084",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 1500kVA 6600/440-254V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 24246000,
        "netPrice": null,
        "costPrice": null
      }
    ]
  },
  {
    "id": "REN-10-085",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 2000kVA 6600/440-254V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 32158000,
        "netPrice": null,
        "costPrice": null
      }
    ]
  },
  {
    "id": "REN-10-088",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 150kVA 6600/550-318V 60Hz 油入形",
    "spec": "日立",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": 6218000,
    "costPrice": 6218000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-090",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 150kVA 6600/550-318V 60Hz 油入形",
    "spec": "岐阜愛知",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": 3800000,
    "costPrice": 3800000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 7815000,
        "netPrice": null,
        "costPrice": 1633200
      }
    ]
  },
  {
    "id": "REN-10-092",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 500kVA 6600/550-318V 60Hz 油入形",
    "spec": "岐阜愛知",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": 4500000,
    "costPrice": 4500000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 7815000,
        "netPrice": null,
        "costPrice": 1633200
      }
    ]
  },
  {
    "id": "REN-10-096",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "防振ゴム(屋内)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-097",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "防振ゴム(屋内)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-098",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "防振ゴム(屋内)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-099",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "防振ゴム(屋内)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-100",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "ダイヤル温度計(本体取付)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-101",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "ダイヤル温度計(本体取付)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-102",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "ダイヤル温度計(本体取付)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-103",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "ダイヤル温度計(盤面取付)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-104",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "ダイヤル温度計(盤面取付)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-107",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 10kVA 6600/210-105V 60Hz ﾓｰﾙﾄﾞ形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-108",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 20kVA 6600/210-105V 60Hz ﾓｰﾙﾄﾞ形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-109",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 30kVA 6600/210-105V 60Hz ﾓｰﾙﾄﾞ形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-110",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 50kVA 6600/210-105V 60Hz ﾓｰﾙﾄﾞ形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-111",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 75kVA 6600/210-105V 60Hz ﾓｰﾙﾄﾞ形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-112",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 100kVA 6600/210-105V 60Hz ﾓｰﾙﾄﾞ形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-113",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 150kVA 6600/210-105V 60Hz ﾓｰﾙﾄﾞ形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 8594000,
    "netPrice": 2565000,
    "costPrice": 2565000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-114",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 200kVA 6600/210-105V 60Hz ﾓｰﾙﾄﾞ形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 10422000,
    "netPrice": 3110000,
    "costPrice": 3110000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-115",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 300kVA 6600/210-105V 60Hz ﾓｰﾙﾄﾞ形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-116",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 500kVA 6600/210-105V 60Hz ﾓｰﾙﾄﾞ形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-119",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 20kVA 6600/210V 60Hz ﾓｰﾙﾄﾞ形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-120",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 30kVA 6600/210V 60Hz ﾓｰﾙﾄﾞ形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-121",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 50kVA 6600/210V 60Hz ﾓｰﾙﾄﾞ形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-122",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 75kVA 6600/210V 60Hz ﾓｰﾙﾄﾞ形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-123",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 100kVA 6600/210V 60Hz ﾓｰﾙﾄﾞ形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-124",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 150kVA 6600/210V 60Hz ﾓｰﾙﾄﾞ形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-125",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 200kVA 6600/210V 60Hz ﾓｰﾙﾄﾞ形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-126",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 300kVA 6600/210V 60Hz ﾓｰﾙﾄﾞ形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 15691000,
    "netPrice": 4683000,
    "costPrice": 4683000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-127",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 500kVA 6600/210V 60Hz ﾓｰﾙﾄﾞ形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 23176000,
    "netPrice": 6918000,
    "costPrice": 6918000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-128",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 750kVA 6600/210V 60Hz ﾓｰﾙﾄﾞ形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-129",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 1000kVA 6600/210V 60Hz ﾓｰﾙﾄﾞ形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-130",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 1500kVA 6600/210V 60Hz ﾓｰﾙﾄﾞ形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-131",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 2000kVA 6600/210V 60Hz ﾓｰﾙﾄﾞ形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-134",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ4W 300kVA 6600/182-105V 60Hz ﾓｰﾙﾄﾞ形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 18988000,
    "netPrice": 5668000,
    "costPrice": 5668000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-135",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "ｽｺｯﾄTr 50kVA 210/210-105V 60Hz ﾓｰﾙﾄﾞ形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 4335000,
    "netPrice": 1294000,
    "costPrice": 1294000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-136",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 40kVA 210/210-105V 60Hz ﾓｰﾙﾄﾞ形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 6075000,
    "netPrice": 1858000,
    "costPrice": 1858000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-139",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "温度計",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 314000,
    "netPrice": 93600,
    "costPrice": 93600,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-140",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "防振ｺﾞﾑ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 110000,
    "netPrice": 32800,
    "costPrice": 32800,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-145",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 10kVA 6600/210-105V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 835000,
    "netPrice": 205530,
    "costPrice": 158100,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 792000,
        "netPrice": null,
        "costPrice": 193000
      }
    ]
  },
  {
    "id": "REN-10-146",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 20kVA 6600/210-105V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1093000,
    "netPrice": 269100,
    "costPrice": 207000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 928000,
        "netPrice": null,
        "costPrice": 225000
      }
    ]
  },
  {
    "id": "REN-10-147",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 30kVA 6600/210-105V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1212000,
    "netPrice": 298350,
    "costPrice": 229500,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1054000,
        "netPrice": null,
        "costPrice": 255500
      }
    ]
  },
  {
    "id": "REN-10-148",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 50kVA 6600/210-105V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1519000,
    "netPrice": 374010,
    "costPrice": 287700,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1294000,
        "netPrice": null,
        "costPrice": 313000
      }
    ]
  },
  {
    "id": "REN-10-149",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 75kVA 6600/210-105V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1834000,
    "netPrice": 451620,
    "costPrice": 347400,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1600000,
        "netPrice": null,
        "costPrice": 388000
      }
    ]
  },
  {
    "id": "REN-10-150",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 100kVA 6600/210-105V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2162000,
    "netPrice": 532350,
    "costPrice": 409500,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1933000,
        "netPrice": null,
        "costPrice": 468600
      }
    ]
  },
  {
    "id": "REN-10-151",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 150kVA 6600/210-105V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2830000,
    "netPrice": 696930,
    "costPrice": 536100,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 2587000,
        "netPrice": null,
        "costPrice": 627200
      }
    ]
  },
  {
    "id": "REN-10-152",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 200kVA 6600/210-105V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3626000,
    "netPrice": 893100,
    "costPrice": 687000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 3199000,
        "netPrice": null,
        "costPrice": 775000
      }
    ]
  },
  {
    "id": "REN-10-153",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 300kVA 6600/210-105V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 4944000,
    "netPrice": 1217580,
    "costPrice": 936600,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 4457000,
        "netPrice": null,
        "costPrice": 1080500
      }
    ]
  },
  {
    "id": "REN-10-154",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 500kVA 6600/210-105V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 7912000,
    "netPrice": 1948830,
    "costPrice": 1499100,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 7310000,
        "netPrice": null,
        "costPrice": 1772800
      }
    ]
  },
  {
    "id": "REN-10-157",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 10kVA 6600/210-105V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 919000,
    "netPrice": 226200,
    "costPrice": 174000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 871000,
        "netPrice": null,
        "costPrice": 211300
      }
    ]
  },
  {
    "id": "REN-10-158",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 20kVA 6600/210-105V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1202000,
    "netPrice": 296010,
    "costPrice": 227700,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1022000,
        "netPrice": null,
        "costPrice": 247600
      }
    ]
  },
  {
    "id": "REN-10-159",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 30kVA 6600/210-105V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1334000,
    "netPrice": 328380,
    "costPrice": 252600,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1160000,
        "netPrice": null,
        "costPrice": 281000
      }
    ]
  },
  {
    "id": "REN-10-160",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 50kVA 6600/210-105V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1671000,
    "netPrice": 411450,
    "costPrice": 316500,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1422000,
        "netPrice": null,
        "costPrice": 344800
      }
    ]
  },
  {
    "id": "REN-10-161",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 75kVA 6600/210-105V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2018000,
    "netPrice": 496860,
    "costPrice": 382200,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1763000,
        "netPrice": null,
        "costPrice": 427300
      }
    ]
  },
  {
    "id": "REN-10-162",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 100kVA 6600/210-105V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2379000,
    "netPrice": 585780,
    "costPrice": 450600,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 2128000,
        "netPrice": null,
        "costPrice": 516000
      }
    ]
  },
  {
    "id": "REN-10-163",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 150kVA 6600/210-105V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3112000,
    "netPrice": 766350,
    "costPrice": 589500,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 2844000,
        "netPrice": null,
        "costPrice": 689600
      }
    ]
  },
  {
    "id": "REN-10-164",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 200kVA 6600/210-105V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3989000,
    "netPrice": 982410,
    "costPrice": 755700,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 3519000,
        "netPrice": null,
        "costPrice": 853400
      }
    ]
  },
  {
    "id": "REN-10-165",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 300kVA 6600/210-105V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 5438000,
    "netPrice": 1339260,
    "costPrice": 1030200,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 4904000,
        "netPrice": null,
        "costPrice": 1189000
      }
    ]
  },
  {
    "id": "REN-10-166",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 500kVA 6600/210-105V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 8704000,
    "netPrice": 2143830,
    "costPrice": 1649100,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 8042000,
        "netPrice": null,
        "costPrice": 1950000
      }
    ]
  },
  {
    "id": "REN-10-169",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 20kVA 6600/210V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1297000,
    "netPrice": 319410,
    "costPrice": 245700,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1146000,
        "netPrice": null,
        "costPrice": 277800
      }
    ]
  },
  {
    "id": "REN-10-170",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 30kVA 6600/210V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1501000,
    "netPrice": 369720,
    "costPrice": 284400,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1260000,
        "netPrice": null,
        "costPrice": 305500
      }
    ]
  },
  {
    "id": "REN-10-171",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 50kVA 6600/210V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1785000,
    "netPrice": 439530,
    "costPrice": 338100,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1512000,
        "netPrice": null,
        "costPrice": 366600
      }
    ]
  },
  {
    "id": "REN-10-172",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 75kVA 6600/210V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2081000,
    "netPrice": 512460,
    "costPrice": 394200,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1808000,
        "netPrice": null,
        "costPrice": 438400
      }
    ]
  },
  {
    "id": "REN-10-173",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 100kVA 6600/210V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2402000,
    "netPrice": 591630,
    "costPrice": 455100,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 2128000,
        "netPrice": null,
        "costPrice": 516000
      }
    ]
  },
  {
    "id": "REN-10-174",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 150kVA 6600/210V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3370000,
    "netPrice": 829920,
    "costPrice": 638400,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 2743000,
        "netPrice": null,
        "costPrice": 665200
      }
    ]
  },
  {
    "id": "REN-10-175",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 200kVA 6600/210V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3941000,
    "netPrice": 970710,
    "costPrice": 746700,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 3338000,
        "netPrice": null,
        "costPrice": 809500
      }
    ]
  },
  {
    "id": "REN-10-176",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 300kVA 6600/210V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 5531000,
    "netPrice": 1362270,
    "costPrice": 1047900,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 4906000,
        "netPrice": null,
        "costPrice": 1189800
      }
    ]
  },
  {
    "id": "REN-10-177",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 500kVA 6600/210V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 8620000,
    "netPrice": 2123160,
    "costPrice": 1633200,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 7815000,
        "netPrice": null,
        "costPrice": 1895000
      }
    ]
  },
  {
    "id": "REN-10-178",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 750kVA 6600/210V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 12710000,
    "netPrice": 3128970,
    "costPrice": 2406900,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 11541000,
        "netPrice": null,
        "costPrice": 2798900
      }
    ]
  },
  {
    "id": "REN-10-179",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 1000kVA 6600/210V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 16310000,
    "netPrice": 4017000,
    "costPrice": 3090000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 15386000,
        "netPrice": null,
        "costPrice": 3731600
      }
    ]
  },
  {
    "id": "REN-10-184",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 20kVA 6600/210V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1425000,
    "netPrice": 351000,
    "costPrice": 270000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1260000,
        "netPrice": null,
        "costPrice": 305500
      }
    ]
  },
  {
    "id": "REN-10-185",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 30kVA 6600/210V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1652000,
    "netPrice": 406770,
    "costPrice": 312900,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1387000,
        "netPrice": null,
        "costPrice": 336300
      }
    ]
  },
  {
    "id": "REN-10-186",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 50kVA 6600/210V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1962000,
    "netPrice": 483210,
    "costPrice": 371700,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1661000,
        "netPrice": null,
        "costPrice": 402600
      }
    ]
  },
  {
    "id": "REN-10-187",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 75kVA 6600/210V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2288000,
    "netPrice": 563550,
    "costPrice": 433500,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1992000,
        "netPrice": null,
        "costPrice": 482900
      }
    ]
  },
  {
    "id": "REN-10-188",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 100kVA 6600/210V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2645000,
    "netPrice": 651300,
    "costPrice": 501000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 2340000,
        "netPrice": null,
        "costPrice": 567400
      }
    ]
  },
  {
    "id": "REN-10-189",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 150kVA 6600/210V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3705000,
    "netPrice": 912600,
    "costPrice": 702000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 3018000,
        "netPrice": null,
        "costPrice": 731900
      }
    ]
  },
  {
    "id": "REN-10-190",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 200kVA 6600/210V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 4334000,
    "netPrice": 1067430,
    "costPrice": 821100,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 3669000,
        "netPrice": null,
        "costPrice": 889800
      }
    ]
  },
  {
    "id": "REN-10-191",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 300kVA 6600/210V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 6080000,
    "netPrice": 1497600,
    "costPrice": 1152000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 5398000,
        "netPrice": null,
        "costPrice": 1309000
      }
    ]
  },
  {
    "id": "REN-10-192",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 500kVA 6600/210V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 9481000,
    "netPrice": 2335320,
    "costPrice": 1796400,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 8598000,
        "netPrice": null,
        "costPrice": 2085000
      }
    ]
  },
  {
    "id": "REN-10-193",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 750kVA 6600/210V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 13980000,
    "netPrice": 3441750,
    "costPrice": 2647500,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 12696000,
        "netPrice": null,
        "costPrice": 3079000
      }
    ]
  },
  {
    "id": "REN-10-194",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 1000kVA 6600/210V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 17940000,
    "netPrice": 4418700,
    "costPrice": 3399000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 16926000,
        "netPrice": null,
        "costPrice": 4105000
      }
    ]
  },
  {
    "id": "REN-10-199",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 20kVA 6600/440-254V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1489000,
    "netPrice": 366600,
    "costPrice": 282000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1374000,
        "netPrice": null,
        "costPrice": 333000
      }
    ]
  },
  {
    "id": "REN-10-200",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 30kVA 6600/440-254V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1725000,
    "netPrice": 424710,
    "costPrice": 326700,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1513000,
        "netPrice": null,
        "costPrice": 366900
      }
    ]
  },
  {
    "id": "REN-10-201",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 50kVA 6600/440-254V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2054000,
    "netPrice": 505830,
    "costPrice": 389100,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1812000,
        "netPrice": null,
        "costPrice": 439400
      }
    ]
  },
  {
    "id": "REN-10-202",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 75kVA 6600/440-254V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2391000,
    "netPrice": 588900,
    "costPrice": 453000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 2079000,
        "netPrice": null,
        "costPrice": 504000
      }
    ]
  },
  {
    "id": "REN-10-203",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 100kVA 6600/440-254V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2763000,
    "netPrice": 680550,
    "costPrice": 523500,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 2447000,
        "netPrice": null,
        "costPrice": 593400
      }
    ]
  },
  {
    "id": "REN-10-204",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 150kVA 6600/440-254V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3875000,
    "netPrice": 954330,
    "costPrice": 734100,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 3154000,
        "netPrice": null,
        "costPrice": 764700
      }
    ]
  },
  {
    "id": "REN-10-205",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 200kVA 6600/440-254V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 4529000,
    "netPrice": 1115400,
    "costPrice": 858000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 3841000,
        "netPrice": null,
        "costPrice": 931400
      }
    ]
  },
  {
    "id": "REN-10-206",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 300kVA 6600/440-254V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 6362000,
    "netPrice": 1567020,
    "costPrice": 1205400,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 5398000,
        "netPrice": null,
        "costPrice": 1309000
      }
    ]
  },
  {
    "id": "REN-10-207",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 500kVA 6600/440-254V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 9912000,
    "netPrice": 2441400,
    "costPrice": 1878000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 8598000,
        "netPrice": null,
        "costPrice": 2085200
      }
    ]
  },
  {
    "id": "REN-10-208",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 750kVA 6600/440-254V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 13980000,
    "netPrice": 3441750,
    "costPrice": 2647500,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 12118000,
        "netPrice": null,
        "costPrice": 2938900
      }
    ]
  },
  {
    "id": "REN-10-209",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 1000kVA 6600/440-254V 60Hz 油入形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 17940000,
    "netPrice": 4418700,
    "costPrice": 3399000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 16157000,
        "netPrice": null,
        "costPrice": 3918500
      }
    ]
  },
  {
    "id": "REN-10-214",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 20kVA 6600/440-254V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1639000,
    "netPrice": 403650,
    "costPrice": 310500,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1512000,
        "netPrice": null,
        "costPrice": 366600
      }
    ]
  },
  {
    "id": "REN-10-215",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 30kVA 6600/440-254V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1899000,
    "netPrice": 467610,
    "costPrice": 359700,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1663000,
        "netPrice": null,
        "costPrice": 403300
      }
    ]
  },
  {
    "id": "REN-10-216",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 50kVA 6600/440-254V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2257000,
    "netPrice": 555750,
    "costPrice": 427500,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1996000,
        "netPrice": null,
        "costPrice": 483900
      }
    ]
  },
  {
    "id": "REN-10-217",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 75kVA 6600/440-254V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 2634000,
    "netPrice": 648570,
    "costPrice": 498900,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 2288000,
        "netPrice": null,
        "costPrice": 554700
      }
    ]
  },
  {
    "id": "REN-10-218",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 100kVA 6600/440-254V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3039000,
    "netPrice": 748410,
    "costPrice": 575700,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 2693000,
        "netPrice": null,
        "costPrice": 652900
      }
    ]
  },
  {
    "id": "REN-10-219",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 150kVA 6600/440-254V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 4263000,
    "netPrice": 1049880,
    "costPrice": 807600,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 3471000,
        "netPrice": null,
        "costPrice": 841700
      }
    ]
  },
  {
    "id": "REN-10-220",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 200kVA 6600/440-254V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 4985000,
    "netPrice": 1227720,
    "costPrice": 944400,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 4224000,
        "netPrice": null,
        "costPrice": 1024000
      }
    ]
  },
  {
    "id": "REN-10-221",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 300kVA 6600/440-254V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 6997000,
    "netPrice": 1723410,
    "costPrice": 1325700,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 5938000,
        "netPrice": null,
        "costPrice": 1440000
      }
    ]
  },
  {
    "id": "REN-10-222",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 500kVA 6600/440-254V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 10910000,
    "netPrice": 2685150,
    "costPrice": 2065500,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 9457000,
        "netPrice": null,
        "costPrice": 2293500
      }
    ]
  },
  {
    "id": "REN-10-223",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 750kVA 6600/440-254V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 15380000,
    "netPrice": 3786120,
    "costPrice": 2912400,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 13331000,
        "netPrice": null,
        "costPrice": 3233000
      }
    ]
  },
  {
    "id": "REN-10-224",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 1000kVA 6600/440-254V 60Hz 油入形 混触防止板付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 19730000,
    "netPrice": 4859400,
    "costPrice": 3738000,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 17773000,
        "netPrice": null,
        "costPrice": 4310400
      }
    ]
  },
  {
    "id": "REN-10-229",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "防振ゴム",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 110000,
    "netPrice": 27000,
    "costPrice": 26910,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 182200,
        "netPrice": null,
        "costPrice": 44000,
        "sourceName": "防振ゴム(屋内)"
      }
    ]
  },
  {
    "id": "REN-10-230",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "防振ゴム",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 194000,
    "netPrice": 48000,
    "costPrice": 47580,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 218400,
        "netPrice": null,
        "costPrice": 53000,
        "sourceName": "防振ゴム(屋内)"
      }
    ]
  },
  {
    "id": "REN-10-231",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "防振ゴム",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 110000,
    "netPrice": 27000,
    "costPrice": 26910,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 254600,
        "netPrice": null,
        "costPrice": 62000,
        "sourceName": "防振ゴム(屋内)"
      }
    ]
  },
  {
    "id": "REN-10-232",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "防振ゴム",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 194000,
    "netPrice": 48000,
    "costPrice": 47580,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 507900,
        "netPrice": null,
        "costPrice": 123000,
        "sourceName": "防振ゴム(屋内)"
      }
    ]
  },
  {
    "id": "REN-10-233",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "防振ゴム",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 330000,
    "netPrice": 82000,
    "costPrice": 81120,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 434200,
        "netPrice": null,
        "costPrice": 105000,
        "sourceName": "ダイヤル温度計(本体取付)"
      }
    ]
  },
  {
    "id": "REN-10-236",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "ダイヤル温度計",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 330000,
    "netPrice": 82000,
    "costPrice": 81120,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 507900,
        "netPrice": null,
        "costPrice": 123000,
        "sourceName": "ダイヤル温度計(盤面取付)"
      }
    ]
  },
  {
    "id": "REN-10-237",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "ダイヤル温度計",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 330000,
    "netPrice": 82000,
    "costPrice": 81120,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 398000,
        "netPrice": null,
        "costPrice": 97000,
        "sourceName": "ダイヤル温度計(盤面取付)"
      }
    ]
  },
  {
    "id": "REN-10-238",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "ダイヤル温度計",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 602000,
    "netPrice": 149000,
    "costPrice": 148200,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-246",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 10kVA 6600/210-105V 60Hz 油入形 ｽｰﾊﾟｰ高効率β",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 821000,
        "netPrice": null,
        "costPrice": 493000,
        "sourceName": "Tr1Φ 10kVA 6600/210-105V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-247",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 20kVA 6600/210-105V 60Hz 油入形 ｽｰﾊﾟｰ高効率β",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 847000,
        "netPrice": null,
        "costPrice": 508000,
        "sourceName": "Tr1Φ 20kVA 6600/210-105V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-248",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 30kVA 6600/210-105V 60Hz 油入形 ｽｰﾊﾟｰ高効率β",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 948000,
        "netPrice": null,
        "costPrice": 569000,
        "sourceName": "Tr1Φ 30kVA 6600/210-105V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-249",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 50kVA 6600/210-105V 60Hz 油入形 ｽｰﾊﾟｰ高効率β",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1243000,
        "netPrice": null,
        "costPrice": 746000,
        "sourceName": "Tr1Φ 50kVA 6600/210-105V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-250",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 75kVA 6600/210-105V 60Hz 油入形 ｽｰﾊﾟｰ高効率β",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1758000,
        "netPrice": null,
        "costPrice": 1054000,
        "sourceName": "Tr1Φ 75kVA 6600/210-105V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-251",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 100kVA 6600/210-105V 60Hz 油入形 ｽｰﾊﾟｰ高効率β",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1986000,
        "netPrice": null,
        "costPrice": 1192000,
        "sourceName": "Tr1Φ 100kVA 6600/210-105V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-252",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 150kVA 6600/210-105V 60Hz 油入形 ｽｰﾊﾟｰ高効率β",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 2272000,
        "netPrice": null,
        "costPrice": 1363000,
        "sourceName": "Tr1Φ 150kVA 6600/210-105V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-253",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 200kVA 6600/210-105V 60Hz 油入形 ｽｰﾊﾟｰ高効率β",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 3020000,
        "netPrice": null,
        "costPrice": 1812000,
        "sourceName": "Tr1Φ 200kVA 6600/210-105V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-254",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 300kVA 6600/210-105V 60Hz 油入形 ｽｰﾊﾟｰ高効率β",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 3846000,
        "netPrice": null,
        "costPrice": 2308000,
        "sourceName": "Tr1Φ 300kVA 6600/210-105V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-255",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 500kVA 6600/210-105V 60Hz 油入形 ｽｰﾊﾟｰ高効率β",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 4636000,
        "netPrice": null,
        "costPrice": 2782000,
        "sourceName": "Tr1Φ 500kVA 6600/210-105V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-258",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 20kVA 6600/210V 60Hz 油入形 ｽｰﾊﾟｰ高効率β",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1318000,
        "netPrice": null,
        "costPrice": 790000,
        "sourceName": "Tr3Φ 20kVA 6600/210V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-259",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 30kVA 6600/210V 60Hz 油入形 ｽｰﾊﾟｰ高効率β",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1378000,
        "netPrice": null,
        "costPrice": 829000,
        "sourceName": "Tr3Φ 30kVA 6600/210V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-260",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 50kVA 6600/210V 60Hz 油入形 ｽｰﾊﾟｰ高効率β",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1607000,
        "netPrice": null,
        "costPrice": 964000,
        "sourceName": "Tr3Φ 50kVA 6600/210V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-261",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 75kVA 6600/210V 60Hz 油入形 ｽｰﾊﾟｰ高効率β",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1870000,
        "netPrice": null,
        "costPrice": 1122000,
        "sourceName": "Tr3Φ 75kVA 6600/210V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-262",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 100kVA 6600/210V 60Hz 油入形 ｽｰﾊﾟｰ高効率β",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": 943250,
    "costPrice": 943250,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 2189000,
        "netPrice": null,
        "costPrice": 1313000,
        "sourceName": "Tr3Φ 100kVA 6600/210V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-263",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 150kVA 6600/210V 60Hz 油入形 ｽｰﾊﾟｰ高効率β",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": 1322750,
    "costPrice": 1322750,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 2504000,
        "netPrice": null,
        "costPrice": 1502000,
        "sourceName": "Tr3Φ 150kVA 6600/210V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-264",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 200kVA 6600/210V 60Hz 油入形 ｽｰﾊﾟｰ高効率β",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 3337000,
        "netPrice": null,
        "costPrice": 2002000,
        "sourceName": "Tr3Φ 200kVA 6600/210V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-265",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 300kVA 6600/210V 60Hz 油入形 ｽｰﾊﾟｰ高効率β",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 4373000,
        "netPrice": null,
        "costPrice": 2624000,
        "sourceName": "Tr3Φ 300kVA 6600/210V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-266",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 500kVA 6600/210V 60Hz 油入形 ｽｰﾊﾟｰ高効率β",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 5159000,
        "netPrice": null,
        "costPrice": 3095000,
        "sourceName": "Tr3Φ 500kVA 6600/210V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-267",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 750kVA 6600/210V 60Hz 油入形 ｽｰﾊﾟｰ高効率β",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 158000000,
    "netPrice": 5487000,
    "costPrice": 5486250,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 6871000,
        "netPrice": null,
        "costPrice": 4122000,
        "sourceName": "Tr3Φ 750kVA 6600/210V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-268",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 1000kVA 6600/210V 60Hz 油入形 ｽｰﾊﾟｰ高効率β",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 8150000,
        "netPrice": null,
        "costPrice": 4890000,
        "sourceName": "Tr3Φ 1000kVA 6600/210V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-271",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 10kVA 6600/210-105V 60Hz 油入形 ｽｰﾊﾟｰ高効率α",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-272",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 20kVA 6600/210-105V 60Hz 油入形 ｽｰﾊﾟｰ高効率α",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-273",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 30kVA 6600/210-105V 60Hz 油入形 ｽｰﾊﾟｰ高効率α",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 1957000,
        "netPrice": null,
        "costPrice": 1174000,
        "sourceName": "Tr3Φ 75kVA 6600/440-254V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-274",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 50kVA 6600/210-105V 60Hz 油入形 ｽｰﾊﾟｰ高効率α",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 2288000,
        "netPrice": null,
        "costPrice": 1373000,
        "sourceName": "Tr3Φ 100kVA 6600/440-254V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-275",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 75kVA 6600/210-105V 60Hz 油入形 ｽｰﾊﾟｰ高効率α",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 2613000,
        "netPrice": null,
        "costPrice": 1568000,
        "sourceName": "Tr3Φ 150kVA 6600/440-254V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-276",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 100kVA 6600/210-105V 60Hz 油入形 ｽｰﾊﾟｰ高効率α",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 3490000,
        "netPrice": null,
        "costPrice": 2094000,
        "sourceName": "Tr3Φ 200kVA 6600/440-254V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-277",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 150kVA 6600/210-105V 60Hz 油入形 ｽｰﾊﾟｰ高効率α",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 4570000,
        "netPrice": null,
        "costPrice": 2742000,
        "sourceName": "Tr3Φ 300kVA 6600/440-254V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-278",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 200kVA 6600/210-105V 60Hz 油入形 ｽｰﾊﾟｰ高効率α",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 5391000,
        "netPrice": null,
        "costPrice": 3235000,
        "sourceName": "Tr3Φ 500kVA 6600/440-254V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-279",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 300kVA 6600/210-105V 60Hz 油入形 ｽｰﾊﾟｰ高効率α",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 7199000,
        "netPrice": null,
        "costPrice": 4319000,
        "sourceName": "Tr3Φ 750kVA 6600/440-254V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-280",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr1Φ 500kVA 6600/210-105V 60Hz 油入形 ｽｰﾊﾟｰ高効率α",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 8543000,
        "netPrice": null,
        "costPrice": 5126000,
        "sourceName": "Tr3Φ 1000kVA 6600/440-254V 60Hz 油入形 Super ｱﾓﾙﾌｧｽ ZeroS"
      }
    ]
  },
  {
    "id": "REN-10-283",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 20kVA 6600/210V 60Hz 油入形 ｽｰﾊﾟｰ高効率α",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-284",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 30kVA 6600/210V 60Hz 油入形 ｽｰﾊﾟｰ高効率α",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-285",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 50kVA 6600/210V 60Hz 油入形 ｽｰﾊﾟｰ高効率α",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 159000,
        "netPrice": null,
        "costPrice": 95000,
        "sourceName": "防振ゴム(屋内)"
      }
    ]
  },
  {
    "id": "REN-10-286",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 75kVA 6600/210V 60Hz 油入形 ｽｰﾊﾟｰ高効率α",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 262000,
        "netPrice": null,
        "costPrice": 157000,
        "sourceName": "防振ゴム(屋内)"
      }
    ]
  },
  {
    "id": "REN-10-287",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 100kVA 6600/210V 60Hz 油入形 ｽｰﾊﾟｰ高効率α",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": 1225950,
    "costPrice": 1225950,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 393000,
        "netPrice": null,
        "costPrice": 236000,
        "sourceName": "防振ゴム(屋内)"
      }
    ]
  },
  {
    "id": "REN-10-288",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 150kVA 6600/210V 60Hz 油入形 ｽｰﾊﾟｰ高効率α",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": 1719850,
    "costPrice": 1719850,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 523000,
        "netPrice": null,
        "costPrice": 314000,
        "sourceName": "防振ゴム(屋内)"
      }
    ]
  },
  {
    "id": "REN-10-289",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 200kVA 6600/210V 60Hz 油入形 ｽｰﾊﾟｰ高効率α",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 315000,
        "netPrice": null,
        "costPrice": 189000,
        "sourceName": "ダイヤル温度計(本体取付)"
      }
    ]
  },
  {
    "id": "REN-10-290",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 300kVA 6600/210V 60Hz 油入形 ｽｰﾊﾟｰ高効率α",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 238000,
        "netPrice": null,
        "costPrice": 143000,
        "sourceName": "ダイヤル温度計(本体取付)"
      }
    ]
  },
  {
    "id": "REN-10-291",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 500kVA 6600/210V 60Hz 油入形 ｽｰﾊﾟｰ高効率α",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-10-292",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 750kVA 6600/210V 60Hz 油入形 ｽｰﾊﾟｰ高効率α",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 368000,
        "netPrice": null,
        "costPrice": 221000,
        "sourceName": "ダイヤル温度計(盤面取付)"
      }
    ]
  },
  {
    "id": "REN-10-293",
    "categoryId": "cat-9",
    "category": "TR",
    "name": "Tr3Φ 1000kVA 6600/210V 60Hz 油入形 ｽｰﾊﾟｰ高効率α",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "日立",
    "effectiveDate": null,
    "history": [
      {
        "label": "TR",
        "effectiveDate": null,
        "listPrice": 289000,
        "netPrice": null,
        "costPrice": 173000,
        "sourceName": "ダイヤル温度計(盤面取付)"
      }
    ]
  },
  {
    "id": "REN-11-003",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "LEG-107L 手動復帰 露出",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 29900,
    "netPrice": 17000,
    "costPrice": 16600,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-004",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "LEG-108L 手動復帰 埋込",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 29900,
    "netPrice": 17000,
    "costPrice": 16600,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-005",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "LEG-109LA 手動/自動復帰 埋込",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 32900,
    "netPrice": 19000,
    "costPrice": 18300,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-008",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "M30",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 16900,
    "netPrice": 11900,
    "costPrice": 11900,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-009",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "SM41",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 22300,
    "netPrice": 15700,
    "costPrice": 15700,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-010",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "SM64",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 24200,
    "netPrice": 17000,
    "costPrice": 17000,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-011",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "SM106",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 45800,
    "netPrice": 32100,
    "costPrice": 32100,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-012",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "SM120",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 80700,
    "netPrice": 56500,
    "costPrice": 56500,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-013",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "SM156BR",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 110800,
    "netPrice": 77600,
    "costPrice": 77600,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-014",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "SM240BR",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 270600,
    "netPrice": 189500,
    "costPrice": 189500,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-015",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "BMC30 旧：BM30",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 10200,
    "netPrice": 7200,
    "costPrice": 7200,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-016",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "BM41 旧：BM41",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 14000,
    "netPrice": 9800,
    "costPrice": 9800,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-017",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "BMC64 旧：BM64",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 20200,
    "netPrice": 14200,
    "costPrice": 14200,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-018",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "BMC106 旧：BM106",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 31700,
    "netPrice": 22200,
    "costPrice": 22200,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-020",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "DM55B",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 65800,
    "netPrice": 46100,
    "costPrice": 46100,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-021",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "DM70B",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 111600,
    "netPrice": 78200,
    "costPrice": 78200,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-022",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "DM100B",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 158800,
    "netPrice": 111200,
    "costPrice": 111200,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-025",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "LD-25 ﾌﾞｻﾞｰ内蔵 露出形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 25800,
    "netPrice": 18000,
    "costPrice": 18000,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-026",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "BM30A 内径30mm 100A 貫通形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 11300,
    "netPrice": 8000,
    "costPrice": 8000,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-027",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "LD-24A ﾌﾞｻﾞｰ外付 埋込形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 40300,
    "netPrice": 25000,
    "costPrice": 24700,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-028",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "BM30A 内径30mm 100A 貫通形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 11300,
    "netPrice": 8000,
    "costPrice": 8000,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-031",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "ELR EF-4A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 10200,
    "netPrice": 7000,
    "costPrice": 7000,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-032",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "ZCT S-20",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 5500,
    "netPrice": 3800,
    "costPrice": 3800,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-033",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "ELR EF-5",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 12200,
    "netPrice": 6060,
    "costPrice": 6060,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-034",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "ZCT S-30",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 11600,
    "netPrice": 5770,
    "costPrice": 5770,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-038",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "NV-ZSA 高速型",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 35800,
    "netPrice": 18000,
    "costPrice": 17010,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-039",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "NV-ZSA 時延型",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 40100,
    "netPrice": 19000,
    "costPrice": 18810,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-041",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "ZCT Φ15",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 9210,
    "netPrice": 5000,
    "costPrice": 4090,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-042",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "ZCT Φ30",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 14900,
    "netPrice": 7000,
    "costPrice": 6620,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-043",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "ZCT Φ40",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 23200,
    "netPrice": 11000,
    "costPrice": 10260,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-044",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "ZCT Φ60",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 63600,
    "netPrice": 29000,
    "costPrice": 28350,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-045",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "ZCT Φ80",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 75100,
    "netPrice": 34000,
    "costPrice": 33390,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-11-046",
    "categoryId": "cat-10",
    "category": "ELR",
    "name": "ZCT Φ100",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 114600,
    "netPrice": 52000,
    "costPrice": 51120,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-003",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "VT 440/110V 15VA",
    "spec": "PE-15F",
    "defaultQty": 1,
    "listPrice": 40000,
    "netPrice": 22000,
    "costPrice": 16000,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-005",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "VT 380/110V 50VA",
    "spec": "PD-50HF",
    "defaultQty": 2,
    "listPrice": 84000,
    "netPrice": 46200,
    "costPrice": 31200,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-006",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "VT 380/110V 15VA",
    "spec": "PE-15F",
    "defaultQty": 2,
    "listPrice": 56000,
    "netPrice": 30800,
    "costPrice": 20740,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-008",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "OPTr 6600/110V 600VA",
    "spec": "EMT-BB",
    "defaultQty": 1,
    "listPrice": 314000,
    "netPrice": 173000,
    "costPrice": 107400,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-012",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 60/5A 5VA 丸窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 11600,
    "netPrice": 7000,
    "costPrice": 4320,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-013",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 75/5A 5VA 丸窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 11600,
    "netPrice": 7000,
    "costPrice": 4320,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-014",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 100/5A 5VA 丸窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 5440,
    "netPrice": 3000,
    "costPrice": 2020,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-015",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 120/5A 5VA 丸窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 5440,
    "netPrice": 3000,
    "costPrice": 2020,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-016",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 150/5A 5VA 丸窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 5800,
    "netPrice": 4000,
    "costPrice": 2150,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-017",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 200/5A 5VA 丸窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 5800,
    "netPrice": 4000,
    "costPrice": 2150,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-018",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 250/5A 5VA 丸窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 6310,
    "netPrice": 4000,
    "costPrice": 2340,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-019",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 300/5A 5VA 丸窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 6310,
    "netPrice": 4000,
    "costPrice": 2340,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-020",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 400/5A 5VA 丸窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 6480,
    "netPrice": 4000,
    "costPrice": 2410,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-021",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 500/5A 5VA 丸窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 8170,
    "netPrice": 5000,
    "costPrice": 3030,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-022",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 600/5A 5VA 丸窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 8580,
    "netPrice": 5000,
    "costPrice": 3190,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-023",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 750/5A 5VA 丸窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 8580,
    "netPrice": 5000,
    "costPrice": 3190,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-025",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 100/5A 15VA 丸窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 12000,
    "netPrice": 7000,
    "costPrice": 4450,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-026",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 120/5A 15VA 丸窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 12000,
    "netPrice": 7000,
    "costPrice": 4450,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-027",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 150/5A 15VA 丸窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 6690,
    "netPrice": 4000,
    "costPrice": 2480,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-028",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 180/5A 15VA 丸窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 7070,
    "netPrice": 4000,
    "costPrice": 2630,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-029",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 200/5A 15VA 丸窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 7070,
    "netPrice": 4000,
    "costPrice": 2630,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-030",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 240/5A 15VA 丸窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 7270,
    "netPrice": 4000,
    "costPrice": 2700,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-031",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 250/5A 15VA 丸窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 7270,
    "netPrice": 4000,
    "costPrice": 2700,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-032",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 300/5A 15VA 丸窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 7270,
    "netPrice": 4000,
    "costPrice": 2700,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-033",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 400/5A 15VA 丸窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 7740,
    "netPrice": 5000,
    "costPrice": 2870,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-034",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 500/5A 15VA 丸窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 8940,
    "netPrice": 5000,
    "costPrice": 3320,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-035",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 600/5A 15VA 丸窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 9400,
    "netPrice": 6000,
    "costPrice": 3500,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-036",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 750/5A 15VA 丸窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 9810,
    "netPrice": 6000,
    "costPrice": 3500,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-038",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 150/5A 15VA 角窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 24000,
    "netPrice": 14000,
    "costPrice": 8910,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-039",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 200/5A 15VA 角窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 9850,
    "netPrice": 6000,
    "costPrice": 3650,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-040",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 250/5A 15VA 角窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 10100,
    "netPrice": 6000,
    "costPrice": 3730,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-041",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 300/5A 15VA 角窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 10300,
    "netPrice": 6000,
    "costPrice": 3820,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-042",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 400/5A 15VA 角窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 10500,
    "netPrice": 6000,
    "costPrice": 3890,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-043",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 500/5A 15VA 角窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 10700,
    "netPrice": 6000,
    "costPrice": 3970,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-044",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 600/5A 15VA 角窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 11900,
    "netPrice": 7000,
    "costPrice": 4430,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-045",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 750/5A 15VA 角窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 12800,
    "netPrice": 8000,
    "costPrice": 4790,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-047",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 200/5A 40VA 角窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 18600,
    "netPrice": 11000,
    "costPrice": 6890,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-048",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 250/5A 40VA 角窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 18900,
    "netPrice": 11000,
    "costPrice": 7020,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-049",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 300/5A 40VA 角窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 19000,
    "netPrice": 11000,
    "costPrice": 7090,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-050",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 400/5A 40VA 角窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 15600,
    "netPrice": 9000,
    "costPrice": 5790,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-051",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 500/5A 40VA 角窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 17900,
    "netPrice": 10000,
    "costPrice": 6630,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-052",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 600/5A 40VA 角窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 19200,
    "netPrice": 11000,
    "costPrice": 7150,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-053",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 750/5A 40VA 角窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 19500,
    "netPrice": 11000,
    "costPrice": 7220,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-054",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 800/5A 40VA 角窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 20000,
    "netPrice": 11000,
    "costPrice": 7410,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-055",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 1000/5A 40VA 角窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 21900,
    "netPrice": 13000,
    "costPrice": 8130,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-056",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 1200/5A 40VA 角窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 22200,
    "netPrice": 13000,
    "costPrice": 8260,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-057",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 1500/5A 40VA 角窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 25200,
    "netPrice": 14000,
    "costPrice": 9360,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-058",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 2000/5A 40VA 角窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 32000,
    "netPrice": 18000,
    "costPrice": 11830,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-059",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 2500/5A 40VA 角窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 42000,
    "netPrice": 24000,
    "costPrice": 15600,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-060",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 3000/5A 40VA 角窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 54200,
    "netPrice": 30000,
    "costPrice": 20150,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-061",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 4000/5A 40VA 角窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 103000,
    "netPrice": 57000,
    "costPrice": 38220,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-062",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 5000/5A 40VA 角窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 189000,
    "netPrice": 104000,
    "costPrice": 70200,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-063",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "CT 6000/5A 40VA 角窓貫通形",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 198000,
    "netPrice": 109000,
    "costPrice": 73450,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-12-066",
    "categoryId": "cat-11",
    "category": "低圧VT・CT",
    "name": "THR",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 7400,
    "netPrice": 4000,
    "costPrice": 3290,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-003",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "AM",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 4550,
    "netPrice": 5000,
    "costPrice": 1660,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-004",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "AM 赤指針付",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 5830,
    "netPrice": 5000,
    "costPrice": 1920,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-005",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "VM",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 5410,
    "netPrice": 5000,
    "costPrice": 1980,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-007",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "AM",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 9990,
    "netPrice": 10000,
    "costPrice": 3650,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-008",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "VM",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 10800,
    "netPrice": 10000,
    "costPrice": 3940,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-009",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "Ｗ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 52900,
    "netPrice": 29000,
    "costPrice": 19330,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-011",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "AM",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 22100,
    "netPrice": 10000,
    "costPrice": 6730,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-012",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "VM",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 22100,
    "netPrice": 10000,
    "costPrice": 6730,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-013",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "AS",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 4000,
    "netPrice": 3000,
    "costPrice": 3000,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-014",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "VS",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 4000,
    "netPrice": 3000,
    "costPrice": 3000,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-015",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "COSΦ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 88600,
    "netPrice": 66400,
    "costPrice": 55000,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-016",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "Ｗ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 61000,
    "netPrice": 20000,
    "costPrice": 11900,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-017",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "VAR",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 69400,
    "netPrice": 52000,
    "costPrice": 43000,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-021",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "AM",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 9420,
    "netPrice": 10000,
    "costPrice": 3440,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-022",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "VM",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 9940,
    "netPrice": 10000,
    "costPrice": 3640,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-023",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "AM",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 17000,
    "netPrice": 10000,
    "costPrice": 10000,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-024",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "VM",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 17900,
    "netPrice": 10000,
    "costPrice": 10000,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-025",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "COSΦ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 87900,
    "netPrice": 35400,
    "costPrice": 32130,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-026",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "Ｗ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 52900,
    "netPrice": 21300,
    "costPrice": 19330,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-028",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "AM",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 8890,
    "netPrice": 10000,
    "costPrice": 3250,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-029",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "VM",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 9420,
    "netPrice": 10000,
    "costPrice": 3440,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-033",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "WHM 3P3W 200V 5A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 57400,
    "netPrice": 27000,
    "costPrice": 26800,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-034",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "特別検定費",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": 14000,
    "costPrice": 13200,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-036",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "WHM 3P3W 200V 5A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 57400,
    "netPrice": 27000,
    "costPrice": 26800,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-037",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "検定専用CT 2000/5A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 28500,
    "netPrice": 16000,
    "costPrice": 11410,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-038",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "検定費",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 41900,
    "netPrice": 27000,
    "costPrice": 26200,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-040",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "WHM 1P3W 100V 5A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 57400,
    "netPrice": 27000,
    "costPrice": 26800,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-041",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "特別検定費",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": 14000,
    "costPrice": 13200,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-043",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "WHM 1P3W 100V 5A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 57400,
    "netPrice": 27000,
    "costPrice": 26800,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-044",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "検定専用CT 1200/5A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 23000,
    "netPrice": 13000,
    "costPrice": 8120,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-045",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "検定費",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 41900,
    "netPrice": 27000,
    "costPrice": 26200,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-048",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "WHM (単相3線式 , 埋込形)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 20700,
    "netPrice": 15000,
    "costPrice": 14500,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-051",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "MDA",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 90700,
    "netPrice": 37000,
    "costPrice": 36260,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-054",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "M2PM-S34R 3P3W 200V 120A 60HZ TEK N 表面",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 73640,
    "netPrice": 40600,
    "costPrice": 34220,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-055",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "M2PM-S34R 1P3W 100V 120A 60HZ TEK N 表面",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 73640,
    "netPrice": 40600,
    "costPrice": 34220,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-056",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "M2PM-S34R 1P3W 100V 30A 60HZ TEK N 表面",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 57970,
    "netPrice": 31900,
    "costPrice": 26970,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-058",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "M1PM-R 1P2W 100V 30A 60HZ TEK N 表面",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 16210,
    "netPrice": 9000,
    "costPrice": 7510,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-060",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "M2PM-S34R 1P3W 100V 5A 60HZ HEN N 表面",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 55700,
    "netPrice": 31900,
    "costPrice": 25970,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-061",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "CW-15LS 200/5A",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 8320,
    "netPrice": 4600,
    "costPrice": 3090,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-062",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "検定費",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 29100,
    "netPrice": 20000,
    "costPrice": 18200,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-064",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "M2PM-S34VR 3P3W 200V 5A 60HZ HEN 埋込",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 61800,
    "netPrice": 34000,
    "costPrice": 32130,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-065",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "CW-15LS 15/5A",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 10500,
    "netPrice": 5800,
    "costPrice": 3890,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-066",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "検定費",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 24000,
    "netPrice": 17000,
    "costPrice": 15000,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-068",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "M2PM-S34VR 1P3W 100V 5A 60HZ HEN N 埋込",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 61800,
    "netPrice": 34000,
    "costPrice": 32130,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-069",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "CW-15LS 50/5A",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 10800,
    "netPrice": 6000,
    "costPrice": 4030,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-070",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "検定費",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 24000,
    "netPrice": 17000,
    "costPrice": 15000,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-072",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "M2PM-S34VR 3P3W 200V 5A 60HZ HEN 埋込",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 61800,
    "netPrice": 34000,
    "costPrice": 32130,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-073",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "CW-15LS 100/5A",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 10800,
    "netPrice": 6000,
    "costPrice": 4030,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-074",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "検定費",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 24000,
    "netPrice": 17000,
    "costPrice": 15000,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-076",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "M2PM-S34VR 3P3W 200V 5A 60HZ HEN 埋込",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 61800,
    "netPrice": 34000,
    "costPrice": 32130,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-077",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "CW-15LS 30/5A",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 10500,
    "netPrice": 5800,
    "costPrice": 3890,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-078",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "検定費",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 24000,
    "netPrice": 17000,
    "costPrice": 15000,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-080",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "M2PM-S34R 1P3W 100V 5A 60HZ HEN N 表面",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 55700,
    "netPrice": 31900,
    "costPrice": 25970,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-081",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "CW-15LMS 300/5A",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 11900,
    "netPrice": 6600,
    "costPrice": 4430,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-082",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "検定費",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 29100,
    "netPrice": 20000,
    "costPrice": 18200,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-084",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "M2PM-S34VR 3P3W 200V 5A 60HZ HEN 埋込",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 61800,
    "netPrice": 34000,
    "costPrice": 32130,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-085",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "CW-15LS 150/5A",
    "spec": "",
    "defaultQty": 2,
    "listPrice": 8320,
    "netPrice": 4600,
    "costPrice": 3090,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-13-086",
    "categoryId": "cat-12",
    "category": "ﾒｰﾀｰ",
    "name": "検定費",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 24000,
    "netPrice": 17000,
    "costPrice": 15000,
    "priceLabel": "定価",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-14-005",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 2P 50A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 15990,
    "netPrice": 4797,
    "costPrice": 2717,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 12300,
        "netPrice": 3700,
        "costPrice": 2090
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 10700,
        "netPrice": 4000,
        "costPrice": 1830
      }
    ]
  },
  {
    "id": "REN-14-006",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 50A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 21710,
    "netPrice": 6513,
    "costPrice": 3692,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 16700,
        "netPrice": 5100,
        "costPrice": 2840
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 14500,
        "netPrice": 5000,
        "costPrice": 2480
      }
    ]
  },
  {
    "id": "REN-14-007",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 60A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 27820,
    "netPrice": 8346,
    "costPrice": 4719,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 21400,
        "netPrice": 6500,
        "costPrice": 3630
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 18600,
        "netPrice": 6000,
        "costPrice": 3170
      }
    ]
  },
  {
    "id": "REN-14-008",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 100A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 49010,
    "netPrice": 14703,
    "costPrice": 8268,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 37700,
        "netPrice": 11400,
        "costPrice": 6360
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 32800,
        "netPrice": 10000,
        "costPrice": 5560
      }
    ]
  },
  {
    "id": "REN-14-009",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 125A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 76050,
    "netPrice": 22815,
    "costPrice": 12883,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 58500,
        "netPrice": 17600,
        "costPrice": 9910
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 50900,
        "netPrice": 16000,
        "costPrice": 8650
      }
    ]
  },
  {
    "id": "REN-14-010",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 225A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 95420,
    "netPrice": 28626,
    "costPrice": 16172,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 73400,
        "netPrice": 22100,
        "costPrice": 12440
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 63800,
        "netPrice": 20000,
        "costPrice": 10880
      }
    ]
  },
  {
    "id": "REN-14-011",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 250A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 133510,
    "netPrice": 40053,
    "costPrice": 23062,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 102700,
        "netPrice": 30900,
        "costPrice": 17740
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 89300,
        "netPrice": 27000,
        "costPrice": 15480
      }
    ]
  },
  {
    "id": "REN-14-012",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 400A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 250510,
    "netPrice": 75153,
    "costPrice": 42770,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 192700,
        "netPrice": 57900,
        "costPrice": 32900
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 167600,
        "netPrice": 51000,
        "costPrice": 28740
      }
    ]
  },
  {
    "id": "REN-14-013",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 600A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 521040,
    "netPrice": 156312,
    "costPrice": 87295,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 400800,
        "netPrice": 120300,
        "costPrice": 67150
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 348500,
        "netPrice": 105000,
        "costPrice": 58680
      }
    ]
  },
  {
    "id": "REN-14-014",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 800A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 771600,
    "netPrice": 231480,
    "costPrice": 130980,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 593500,
        "netPrice": 178100,
        "costPrice": 100760
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 516100,
        "netPrice": 155000,
        "costPrice": 88050
      }
    ]
  },
  {
    "id": "REN-14-016",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 50A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 28730,
    "netPrice": 8619,
    "costPrice": 4862,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 22100,
        "netPrice": 6700,
        "costPrice": 3740
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 19200,
        "netPrice": 6000,
        "costPrice": 3620
      }
    ]
  },
  {
    "id": "REN-14-017",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 60A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 36140,
    "netPrice": 10842,
    "costPrice": 6214,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 27800,
        "netPrice": 8400,
        "costPrice": 4780
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 24200,
        "netPrice": 8000,
        "costPrice": 4620
      }
    ]
  },
  {
    "id": "REN-14-018",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 100A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 70980,
    "netPrice": 21294,
    "costPrice": 11960,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 54600,
        "netPrice": 16400,
        "costPrice": 9200
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 47500,
        "netPrice": 15000,
        "costPrice": 8930
      }
    ]
  },
  {
    "id": "REN-14-019",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 125A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 104780,
    "netPrice": 31434,
    "costPrice": 17771,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 80600,
        "netPrice": 24200,
        "costPrice": 13670
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 70100,
        "netPrice": 21000,
        "costPrice": 13240
      }
    ]
  },
  {
    "id": "REN-14-020",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 225A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 130260,
    "netPrice": 39078,
    "costPrice": 31109,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 100200,
        "netPrice": 30100,
        "costPrice": 23930
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 87100,
        "netPrice": 27000,
        "costPrice": 16630
      }
    ]
  },
  {
    "id": "REN-14-021",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 250A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 182260,
    "netPrice": 54678,
    "costPrice": 43615,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 140200,
        "netPrice": 42100,
        "costPrice": 33550
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 121900,
        "netPrice": 37000,
        "costPrice": 23320
      }
    ]
  },
  {
    "id": "REN-14-022",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 400A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 294060,
    "netPrice": 88218,
    "costPrice": 49348,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 226200,
        "netPrice": 67900,
        "costPrice": 37960
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 196700,
        "netPrice": 59000,
        "costPrice": 36780
      }
    ]
  },
  {
    "id": "REN-14-023",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 600A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 578240,
    "netPrice": 173472,
    "costPrice": 98956,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 444800,
        "netPrice": 133500,
        "costPrice": 76120
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 386800,
        "netPrice": 117000,
        "costPrice": 73740
      }
    ]
  },
  {
    "id": "REN-14-024",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 800A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 970840,
    "netPrice": 291252,
    "costPrice": 166062,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 746800,
        "netPrice": 224100,
        "costPrice": 127740
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 649400,
        "netPrice": 215000,
        "costPrice": 123730
      }
    ]
  },
  {
    "id": "REN-14-026",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 50A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 37050,
    "netPrice": 11115,
    "costPrice": 6318,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 28500,
        "netPrice": 8600,
        "costPrice": 4860
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 24800,
        "netPrice": 8000,
        "costPrice": 4710
      }
    ]
  },
  {
    "id": "REN-14-027",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 60A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 48100,
    "netPrice": 14430,
    "costPrice": 8216,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 37000,
        "netPrice": 11100,
        "costPrice": 6320
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 32200,
        "netPrice": 10000,
        "costPrice": 6120
      }
    ]
  },
  {
    "id": "REN-14-028",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 100A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 97760,
    "netPrice": 29328,
    "costPrice": 17147,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 75200,
        "netPrice": 22600,
        "costPrice": 13190
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 65400,
        "netPrice": 20000,
        "costPrice": 12760
      }
    ]
  },
  {
    "id": "REN-14-029",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 125A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 127140,
    "netPrice": 38142,
    "costPrice": 21879,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 97800,
        "netPrice": 29400,
        "costPrice": 16830
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 85000,
        "netPrice": 26000,
        "costPrice": 16280
      }
    ]
  },
  {
    "id": "REN-14-030",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 225A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 137280,
    "netPrice": 41184,
    "costPrice": 32682,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 105600,
        "netPrice": 31700,
        "costPrice": 25140
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 91800,
        "netPrice": 28000,
        "costPrice": 17470
      }
    ]
  },
  {
    "id": "REN-14-031",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 250A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 192010,
    "netPrice": 57603,
    "costPrice": 45903,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 147700,
        "netPrice": 44400,
        "costPrice": 35310
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 128400,
        "netPrice": 39000,
        "costPrice": 24550
      }
    ]
  },
  {
    "id": "REN-14-032",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 400A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 438100,
    "netPrice": 131430,
    "costPrice": 73333,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 337000,
        "netPrice": 101100,
        "costPrice": 56410
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 293000,
        "netPrice": 88000,
        "costPrice": 54650
      }
    ]
  },
  {
    "id": "REN-14-033",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 600A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 858390,
    "netPrice": 257517,
    "costPrice": 147329,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 660300,
        "netPrice": 198100,
        "costPrice": 113330
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 574200,
        "netPrice": 173000,
        "costPrice": 109780
      }
    ]
  },
  {
    "id": "REN-14-034",
    "categoryId": "cat-13",
    "category": "MCB(表)",
    "name": "MCB 3P 800A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "1.3",
    "effectiveDate": null,
    "history": [
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 876500,
        "netPrice": 263000,
        "costPrice": 167290
      }
    ]
  },
  {
    "id": "REN-15-005",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 2P 50A以下 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 19539,
    "netPrice": 5862,
    "costPrice": 3419,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 15030,
        "netPrice": 4600,
        "costPrice": 2630
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 13070,
        "netPrice": 4000,
        "costPrice": 2550,
        "sourceName": "MCB 2P 50A以下"
      }
    ]
  },
  {
    "id": "REN-15-006",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 50A以下 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 27053,
    "netPrice": 8116,
    "costPrice": 4745,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 20810,
        "netPrice": 6300,
        "costPrice": 3650
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 18070,
        "netPrice": 6000,
        "costPrice": 3530,
        "sourceName": "MCB 3P 50A以下"
      }
    ]
  },
  {
    "id": "REN-15-007",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 60A 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 33163,
    "netPrice": 9949,
    "costPrice": 5772,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 25510,
        "netPrice": 7700,
        "costPrice": 4440
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 22170,
        "netPrice": 7000,
        "costPrice": 4300,
        "sourceName": "MCB 3P 60A"
      }
    ]
  },
  {
    "id": "REN-15-008",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 100A以下 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 61295,
    "netPrice": 18389,
    "costPrice": 10673,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 47150,
        "netPrice": 14200,
        "costPrice": 8210
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 41020,
        "netPrice": 13000,
        "costPrice": 7950,
        "sourceName": "MCB 3P 100A以下"
      }
    ]
  },
  {
    "id": "REN-15-009",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 125A 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 88335,
    "netPrice": 26501,
    "costPrice": 15288,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 67950,
        "netPrice": 20400,
        "costPrice": 11760
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 59120,
        "netPrice": 18000,
        "costPrice": 11380,
        "sourceName": "MCB 3P 125A"
      }
    ]
  },
  {
    "id": "REN-15-010",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 225A以下 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 114660,
    "netPrice": 34398,
    "costPrice": 19981,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 88200,
        "netPrice": 26500,
        "costPrice": 15370
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 76700,
        "netPrice": 23000,
        "costPrice": 14890,
        "sourceName": "MCB 3P 225A以下"
      }
    ]
  },
  {
    "id": "REN-15-011",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 250A 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 152750,
    "netPrice": 45825,
    "costPrice": 26858,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 117500,
        "netPrice": 35300,
        "costPrice": 20660
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 102200,
        "netPrice": 31000,
        "costPrice": 19990,
        "sourceName": "MCB 3P 250A"
      }
    ]
  },
  {
    "id": "REN-15-012",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 400A 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 288860,
    "netPrice": 86658,
    "costPrice": 49504,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 222200,
        "netPrice": 66700,
        "costPrice": 38080
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 193200,
        "netPrice": 58000,
        "costPrice": 36870,
        "sourceName": "MCB 3P 400A"
      }
    ]
  },
  {
    "id": "REN-15-013",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 600A 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 580060,
    "netPrice": 174018,
    "costPrice": 98176,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 446200,
        "netPrice": 133900,
        "costPrice": 75520
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 388000,
        "netPrice": 117000,
        "costPrice": 73130,
        "sourceName": "MCB 3P 600A"
      }
    ]
  },
  {
    "id": "REN-15-014",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 800A 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 832650,
    "netPrice": 249795,
    "costPrice": 141986,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 640500,
        "netPrice": 192200,
        "costPrice": 109220
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 557000,
        "netPrice": 168000,
        "costPrice": 105780,
        "sourceName": "MCB 3P 800A"
      }
    ]
  },
  {
    "id": "REN-15-016",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 50A以下 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 34073,
    "netPrice": 10222,
    "costPrice": 5915,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 26210,
        "netPrice": 7900,
        "costPrice": 4550
      }
    ]
  },
  {
    "id": "REN-15-017",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 60A 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 41483,
    "netPrice": 12445,
    "costPrice": 7267,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 31910,
        "netPrice": 9600,
        "costPrice": 5590
      }
    ]
  },
  {
    "id": "REN-15-018",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 100A以下 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 83265,
    "netPrice": 24980,
    "costPrice": 14365,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 64050,
        "netPrice": 19300,
        "costPrice": 11050
      }
    ]
  },
  {
    "id": "REN-15-019",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 125A 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 117065,
    "netPrice": 35120,
    "costPrice": 20176,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 90050,
        "netPrice": 27100,
        "costPrice": 15520
      }
    ]
  },
  {
    "id": "REN-15-020",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 225A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 149500,
    "netPrice": 44850,
    "costPrice": 26143,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 115000,
        "netPrice": 34500,
        "costPrice": 20110
      }
    ]
  },
  {
    "id": "REN-15-021",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 250A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 201500,
    "netPrice": 60450,
    "costPrice": 35126,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 155000,
        "netPrice": 46500,
        "costPrice": 27020
      }
    ]
  },
  {
    "id": "REN-15-022",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 400A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 332280,
    "netPrice": 99684,
    "costPrice": 56173,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 255600,
        "netPrice": 76700,
        "costPrice": 43210
      }
    ]
  },
  {
    "id": "REN-15-023",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 600A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-15-024",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 800A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-15-026",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 2P 50A以下 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 31239,
    "netPrice": 9372,
    "costPrice": 5382,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 24030,
        "netPrice": 7300,
        "costPrice": 4140
      }
    ]
  },
  {
    "id": "REN-15-027",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 50A以下 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 42393,
    "netPrice": 12718,
    "costPrice": 7371,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 32610,
        "netPrice": 9800,
        "costPrice": 5670
      }
    ]
  },
  {
    "id": "REN-15-028",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 60A 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 53443,
    "netPrice": 16033,
    "costPrice": 9269,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 41110,
        "netPrice": 12400,
        "costPrice": 7130
      }
    ]
  },
  {
    "id": "REN-15-029",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 100A以下 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 110045,
    "netPrice": 33014,
    "costPrice": 19552,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 84650,
        "netPrice": 25400,
        "costPrice": 15040
      }
    ]
  },
  {
    "id": "REN-15-030",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 125A 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 139425,
    "netPrice": 41828,
    "costPrice": 24284,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 107250,
        "netPrice": 32200,
        "costPrice": 18680
      }
    ]
  },
  {
    "id": "REN-15-031",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 225A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 156520,
    "netPrice": 46956,
    "costPrice": 27274,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 120400,
        "netPrice": 36200,
        "costPrice": 20980
      }
    ]
  },
  {
    "id": "REN-15-032",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 250A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-15-033",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 400A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-15-034",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 600A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-15-035",
    "categoryId": "cat-14",
    "category": "MCB(裏) ",
    "name": "MCB 3P 800A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-16-005",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 2P 50A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 22932,
    "netPrice": 6880,
    "costPrice": 3874,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 17640,
        "netPrice": 5300,
        "costPrice": 2980
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 15340,
        "netPrice": 5000,
        "costPrice": 2890
      }
    ]
  },
  {
    "id": "REN-16-006",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 50A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 30589,
    "netPrice": 9177,
    "costPrice": 5109,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 23530,
        "netPrice": 7100,
        "costPrice": 3930
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 20440,
        "netPrice": 7000,
        "costPrice": 3800
      }
    ]
  },
  {
    "id": "REN-16-007",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 60A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 36699,
    "netPrice": 11010,
    "costPrice": 6136,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 28230,
        "netPrice": 8500,
        "costPrice": 4720
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 24540,
        "netPrice": 8000,
        "costPrice": 4570
      }
    ]
  },
  {
    "id": "REN-16-008",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 100A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 66950,
    "netPrice": 20085,
    "costPrice": 11583,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 51500,
        "netPrice": 15500,
        "costPrice": 8910
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 44800,
        "netPrice": 14000,
        "costPrice": 8630
      }
    ]
  },
  {
    "id": "REN-16-009",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 125A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 93990,
    "netPrice": 28197,
    "costPrice": 16198,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 72300,
        "netPrice": 21700,
        "costPrice": 12460
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 62900,
        "netPrice": 19000,
        "costPrice": 12060
      }
    ]
  },
  {
    "id": "REN-16-010",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 225A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 124410,
    "netPrice": 37323,
    "costPrice": 21411,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 95700,
        "netPrice": 28800,
        "costPrice": 16470
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 83200,
        "netPrice": 25000,
        "costPrice": 15980
      }
    ]
  },
  {
    "id": "REN-16-011",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 250A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 162500,
    "netPrice": 48750,
    "costPrice": 28288,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 125000,
        "netPrice": 37500,
        "costPrice": 21760
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 108700,
        "netPrice": 33000,
        "costPrice": 21080
      }
    ]
  },
  {
    "id": "REN-16-012",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 400A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 327730,
    "netPrice": 98319,
    "costPrice": 56589,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 252100,
        "netPrice": 75700,
        "costPrice": 43530
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 219200,
        "netPrice": 66000,
        "costPrice": 42150
      }
    ]
  },
  {
    "id": "REN-16-013",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 600A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 618930,
    "netPrice": 185679,
    "costPrice": 105274,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 476100,
        "netPrice": 142900,
        "costPrice": 80980
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 414000,
        "netPrice": 125000,
        "costPrice": 78410
      }
    ]
  },
  {
    "id": "REN-16-014",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 800A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 873730,
    "netPrice": 262119,
    "costPrice": 149578,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 672100,
        "netPrice": 201700,
        "costPrice": 115060
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 584500,
        "netPrice": 176000,
        "costPrice": 111450
      }
    ]
  },
  {
    "id": "REN-16-016",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 50A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 37609,
    "netPrice": 11283,
    "costPrice": 6279,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 28930,
        "netPrice": 8700,
        "costPrice": 4830
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 25140,
        "netPrice": 8000,
        "costPrice": 4680
      }
    ]
  },
  {
    "id": "REN-16-017",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 60A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 45019,
    "netPrice": 13506,
    "costPrice": 7631,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 34630,
        "netPrice": 10400,
        "costPrice": 5870
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 30140,
        "netPrice": 9000,
        "costPrice": 5680
      }
    ]
  },
  {
    "id": "REN-16-018",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 100A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 88920,
    "netPrice": 26676,
    "costPrice": 15275,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 68400,
        "netPrice": 20600,
        "costPrice": 11750
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 59500,
        "netPrice": 18000,
        "costPrice": 11400
      }
    ]
  },
  {
    "id": "REN-16-019",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 125A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 122720,
    "netPrice": 36816,
    "costPrice": 21073,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 94400,
        "netPrice": 28400,
        "costPrice": 16210
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 82100,
        "netPrice": 25000,
        "costPrice": 15710
      }
    ]
  },
  {
    "id": "REN-16-020",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 225A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 159250,
    "netPrice": 47775,
    "costPrice": 27573,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 122500,
        "netPrice": 36800,
        "costPrice": 21210
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 106500,
        "netPrice": 32000,
        "costPrice": 20550
      }
    ]
  },
  {
    "id": "REN-16-021",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 250A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 211250,
    "netPrice": 63375,
    "costPrice": 36556,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 162500,
        "netPrice": 48800,
        "costPrice": 28120
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 141300,
        "netPrice": 43000,
        "costPrice": 27240
      }
    ]
  },
  {
    "id": "REN-16-022",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 400A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 371150,
    "netPrice": 111345,
    "costPrice": 63258,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 285500,
        "netPrice": 85700,
        "costPrice": 48660
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 248200,
        "netPrice": 75000,
        "costPrice": 47120
      }
    ]
  },
  {
    "id": "REN-16-023",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 600A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 676130,
    "netPrice": 202839,
    "costPrice": 116922,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 520100,
        "netPrice": 156100,
        "costPrice": 89940
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 452300,
        "netPrice": 136000,
        "costPrice": 87120
      }
    ]
  },
  {
    "id": "REN-16-024",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 800A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1073020,
    "netPrice": 321906,
    "costPrice": 184652,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 825400,
        "netPrice": 247700,
        "costPrice": 142040
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 717800,
        "netPrice": 216000,
        "costPrice": 137590
      }
    ]
  },
  {
    "id": "REN-16-026",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 50A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 45929,
    "netPrice": 13779,
    "costPrice": 7735,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 35330,
        "netPrice": 10600,
        "costPrice": 5950
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 30740,
        "netPrice": 10000,
        "costPrice": 5760
      }
    ]
  },
  {
    "id": "REN-16-027",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 60A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 56979,
    "netPrice": 17094,
    "costPrice": 9633,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 43830,
        "netPrice": 13200,
        "costPrice": 7410
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 38140,
        "netPrice": 12000,
        "costPrice": 7170
      }
    ]
  },
  {
    "id": "REN-16-028",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 100A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 115700,
    "netPrice": 34710,
    "costPrice": 20462,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 89000,
        "netPrice": 26700,
        "costPrice": 15740
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 77400,
        "netPrice": 24000,
        "costPrice": 15230
      }
    ]
  },
  {
    "id": "REN-16-029",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 125A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 145080,
    "netPrice": 43524,
    "costPrice": 25181,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 111600,
        "netPrice": 33500,
        "costPrice": 19370
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 97000,
        "netPrice": 30000,
        "costPrice": 18750
      }
    ]
  },
  {
    "id": "REN-16-030",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 225A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 167180,
    "netPrice": 50154,
    "costPrice": 28990,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 128600,
        "netPrice": 38600,
        "costPrice": 22300
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 111800,
        "netPrice": 34000,
        "costPrice": 21530
      }
    ]
  },
  {
    "id": "REN-16-031",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 250A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 221910,
    "netPrice": 66573,
    "costPrice": 38415,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 170700,
        "netPrice": 51300,
        "costPrice": 29550
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 148400,
        "netPrice": 45000,
        "costPrice": 28620
      }
    ]
  },
  {
    "id": "REN-16-032",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 400A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 524030,
    "netPrice": 157209,
    "costPrice": 88933,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 403100,
        "netPrice": 121000,
        "costPrice": 68410
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 350500,
        "netPrice": 106000,
        "costPrice": 66260
      }
    ]
  },
  {
    "id": "REN-16-033",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 600A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 970970,
    "netPrice": 291291,
    "costPrice": 167869,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 746900,
        "netPrice": 224100,
        "costPrice": 129130
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 649500,
        "netPrice": 195000,
        "costPrice": 125050
      }
    ]
  },
  {
    "id": "REN-16-034",
    "categoryId": "cat-15",
    "category": "MCB(FP)",
    "name": "MCB 3P 800A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 959000,
        "netPrice": 288000,
        "costPrice": 182950
      }
    ]
  },
  {
    "id": "REN-17-005",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 2P 50A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 22900,
        "netPrice": 7000,
        "costPrice": 5240
      }
    ]
  },
  {
    "id": "REN-17-006",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 50A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 36010,
    "netPrice": 10803,
    "costPrice": 8112,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 27700,
        "netPrice": 8400,
        "costPrice": 6240
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 24100,
        "netPrice": 8000,
        "costPrice": 5480
      }
    ]
  },
  {
    "id": "REN-17-007",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 60A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 41470,
    "netPrice": 12441,
    "costPrice": 9295,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 31900,
        "netPrice": 9600,
        "costPrice": 7150
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 27700,
        "netPrice": 9000,
        "costPrice": 6280
      }
    ]
  },
  {
    "id": "REN-17-008",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 100A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 58630,
    "netPrice": 17589,
    "costPrice": 13091,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 45100,
        "netPrice": 13530,
        "costPrice": 10070
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 39200,
        "netPrice": 12000,
        "costPrice": 8880
      }
    ]
  },
  {
    "id": "REN-17-009",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 125A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 88920,
    "netPrice": 26676,
    "costPrice": 19929,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 68400,
        "netPrice": 20600,
        "costPrice": 15330
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 59500,
        "netPrice": 18000,
        "costPrice": 13480
      }
    ]
  },
  {
    "id": "REN-17-010",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 225A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 110500,
    "netPrice": 33150,
    "costPrice": 24856,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 85000,
        "netPrice": 25500,
        "costPrice": 19120
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 73900,
        "netPrice": 23000,
        "costPrice": 16840
      }
    ]
  },
  {
    "id": "REN-17-011",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 250A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 221000,
    "netPrice": 66300,
    "costPrice": 49608,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 170000,
        "netPrice": 51000,
        "costPrice": 38160
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 147800,
        "netPrice": 45000,
        "costPrice": 33600
      }
    ]
  },
  {
    "id": "REN-17-012",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 400A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 299910,
    "netPrice": 89973,
    "costPrice": 55926,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 230700,
        "netPrice": 69300,
        "costPrice": 43020
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 200600,
        "netPrice": 61000,
        "costPrice": 37880
      }
    ]
  },
  {
    "id": "REN-17-013",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 600A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 659100,
    "netPrice": 197730,
    "costPrice": 122369,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 507000,
        "netPrice": 152100,
        "costPrice": 94130
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 440900,
        "netPrice": 133000,
        "costPrice": 82880
      }
    ]
  },
  {
    "id": "REN-17-015",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 50A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-17-016",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 60A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-17-017",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 100A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-17-018",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 125A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-17-019",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 225A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 276250,
    "netPrice": 82875,
    "costPrice": 50934,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 212500,
        "netPrice": 64000,
        "costPrice": 39180
      }
    ]
  },
  {
    "id": "REN-17-020",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 250A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-17-021",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 400A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-17-022",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 600A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-17-023",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 800A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1122940,
    "netPrice": 336882,
    "costPrice": 208949,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 863800,
        "netPrice": 259140,
        "costPrice": 160730
      }
    ]
  },
  {
    "id": "REN-17-025",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 50A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-17-026",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 60A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-17-027",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 100A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-17-028",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 125A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-17-029",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 225A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 290030,
    "netPrice": 87009,
    "costPrice": 53560,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 223100,
        "netPrice": 67000,
        "costPrice": 41200
      }
    ]
  },
  {
    "id": "REN-17-030",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 250A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-17-031",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 400A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-17-032",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 600A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-17-033",
    "categoryId": "cat-16",
    "category": "ELB(表)",
    "name": "ELB 3P 800A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-18-005",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 2P 50A以下 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-18-006",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 50A以下 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-18-007",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 60A 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-18-008",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 100A以下 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-18-009",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 125A 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-18-010",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 225A以下 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-18-011",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 250A 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-18-012",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 400A 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 338260,
    "netPrice": 101478,
    "costPrice": 62595,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 260200,
        "netPrice": 78100,
        "costPrice": 48150
      }
    ]
  },
  {
    "id": "REN-18-013",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 600A 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-18-015",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 50A以下 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-18-016",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 60A 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-18-017",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 100A以下 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-18-018",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 125A 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-18-019",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 225A以下 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-18-020",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 250A 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-18-021",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 400A 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 555360,
    "netPrice": 166608,
    "costPrice": 102648,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 427200,
        "netPrice": 128200,
        "costPrice": 78960
      }
    ]
  },
  {
    "id": "REN-18-022",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 600A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-18-023",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 800A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-18-025",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 50A以下 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 63050,
    "netPrice": 18915,
    "costPrice": 11713,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 48500,
        "netPrice": 14600,
        "costPrice": 9010
      }
    ]
  },
  {
    "id": "REN-18-026",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 60A 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-18-027",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 100A以下 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-18-028",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 125A 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-18-029",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 225A以下 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 309270,
    "netPrice": 92781,
    "costPrice": 57369,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 237900,
        "netPrice": 71400,
        "costPrice": 44130
      }
    ]
  },
  {
    "id": "REN-18-030",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 250A 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-18-031",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 400A 裏面形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-18-032",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 600A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-18-033",
    "categoryId": "cat-17",
    "category": "ELB(裏)",
    "name": "ELB 3P 800A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": []
  },
  {
    "id": "REN-19-005",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 2P 50A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 27540,
        "netPrice": 9000,
        "costPrice": 6620
      }
    ]
  },
  {
    "id": "REN-19-006",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 50A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 44889,
    "netPrice": 13467,
    "costPrice": 9529,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 34530,
        "netPrice": 10400,
        "costPrice": 7330
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 30040,
        "netPrice": 9000,
        "costPrice": 7080
      }
    ]
  },
  {
    "id": "REN-19-007",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 60A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 50349,
    "netPrice": 15105,
    "costPrice": 10712,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 38730,
        "netPrice": 11700,
        "costPrice": 8240
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 33640,
        "netPrice": 10000,
        "costPrice": 7960
      }
    ]
  },
  {
    "id": "REN-19-008",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 100A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 76570,
    "netPrice": 22971,
    "costPrice": 16406,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 58900,
        "netPrice": 17700,
        "costPrice": 12620
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 51200,
        "netPrice": 16000,
        "costPrice": 12240
      }
    ]
  },
  {
    "id": "REN-19-009",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 125A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 106860,
    "netPrice": 32058,
    "costPrice": 23231,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 82200,
        "netPrice": 24700,
        "costPrice": 17870
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 71500,
        "netPrice": 22000,
        "costPrice": 17300
      }
    ]
  },
  {
    "id": "REN-19-010",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 225A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 139490,
    "netPrice": 41847,
    "costPrice": 30095,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 107300,
        "netPrice": 32200,
        "costPrice": 23150
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 93300,
        "netPrice": 28000,
        "costPrice": 22440
      }
    ]
  },
  {
    "id": "REN-19-011",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 250A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 249990,
    "netPrice": 74997,
    "costPrice": 546247,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 192300,
        "netPrice": 57700,
        "costPrice": 420190
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 167200,
        "netPrice": 51000,
        "costPrice": 40880
      }
    ]
  },
  {
    "id": "REN-19-012",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 400A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 377130,
    "netPrice": 113139,
    "costPrice": 69680,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 290100,
        "netPrice": 87100,
        "costPrice": 53600
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 252200,
        "netPrice": 76000,
        "costPrice": 51920
      }
    ]
  },
  {
    "id": "REN-19-013",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 600A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 756990,
    "netPrice": 227097,
    "costPrice": 140192,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 582300,
        "netPrice": 174700,
        "costPrice": 107840
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 506400,
        "netPrice": 152000,
        "costPrice": 104460
      }
    ]
  },
  {
    "id": "REN-19-015",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 50A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 63869,
    "netPrice": 19161,
    "costPrice": 11583,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 49130,
        "netPrice": 14800,
        "costPrice": 8910
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 42740,
        "netPrice": 13000,
        "costPrice": 8620
      }
    ]
  },
  {
    "id": "REN-19-016",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 60A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 77779,
    "netPrice": 23334,
    "costPrice": 14105,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 59830,
        "netPrice": 18000,
        "costPrice": 10850
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 52040,
        "netPrice": 16000,
        "costPrice": 10520
      }
    ]
  },
  {
    "id": "REN-19-017",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 100A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 211380,
    "netPrice": 63414,
    "costPrice": 39156,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 162600,
        "netPrice": 48800,
        "costPrice": 30120
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 141400,
        "netPrice": 43000,
        "costPrice": 29180
      }
    ]
  },
  {
    "id": "REN-19-018",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 125A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 238940,
    "netPrice": 71682,
    "costPrice": 43979,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 183800,
        "netPrice": 55200,
        "costPrice": 33830
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 159800,
        "netPrice": 48000,
        "costPrice": 32780
      }
    ]
  },
  {
    "id": "REN-19-019",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 225A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 305240,
    "netPrice": 91572,
    "costPrice": 56173,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 234800,
        "netPrice": 70500,
        "costPrice": 43210
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 204200,
        "netPrice": 62000,
        "costPrice": 41890
      }
    ]
  },
  {
    "id": "REN-19-020",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 250A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 415740,
    "netPrice": 124722,
    "costPrice": 76713,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 319800,
        "netPrice": 96000,
        "costPrice": 59010
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 278100,
        "netPrice": 84000,
        "costPrice": 57160
      }
    ]
  },
  {
    "id": "REN-19-021",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 400A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 594230,
    "netPrice": 178269,
    "costPrice": 109733,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 457100,
        "netPrice": 137200,
        "costPrice": 84410
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 397400,
        "netPrice": 119300,
        "costPrice": 81750
      }
    ]
  },
  {
    "id": "REN-19-022",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 600A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 859690,
    "netPrice": 257907,
    "costPrice": 158158,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 661300,
        "netPrice": 198400,
        "costPrice": 121660
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 575100,
        "netPrice": 172600,
        "costPrice": 117830
      }
    ]
  },
  {
    "id": "REN-19-023",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 800A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1225120,
    "netPrice": 367536,
    "costPrice": 227526,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 942400,
        "netPrice": 282800,
        "costPrice": 175020
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 819500,
        "netPrice": 245900,
        "costPrice": 169530
      }
    ]
  },
  {
    "id": "REN-19-025",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 50A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 71929,
    "netPrice": 21579,
    "costPrice": 13130,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 55330,
        "netPrice": 16600,
        "costPrice": 10100
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 48140,
        "netPrice": 15000,
        "costPrice": 9770
      }
    ]
  },
  {
    "id": "REN-19-026",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 60A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 88180,
    "netPrice": 26454,
    "costPrice": 16003,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 67831,
        "netPrice": 20400,
        "costPrice": 12310
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 58940,
        "netPrice": 18000,
        "costPrice": 11920
      }
    ]
  },
  {
    "id": "REN-19-027",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 100A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 238940,
    "netPrice": 71682,
    "costPrice": 43979,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 183800,
        "netPrice": 55200,
        "costPrice": 33830
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 159800,
        "netPrice": 48000,
        "costPrice": 32780
      }
    ]
  },
  {
    "id": "REN-19-028",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 125A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 256230,
    "netPrice": 76869,
    "costPrice": 47268,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 197100,
        "netPrice": 59200,
        "costPrice": 36360
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 171400,
        "netPrice": 52000,
        "costPrice": 35200
      }
    ]
  },
  {
    "id": "REN-19-029",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 225A以下",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 319930,
    "netPrice": 95979,
    "costPrice": 59007,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 246100,
        "netPrice": 73900,
        "costPrice": 45390
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 214000,
        "netPrice": 65000,
        "costPrice": 43970
      }
    ]
  },
  {
    "id": "REN-19-030",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 250A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 436150,
    "netPrice": 130845,
    "costPrice": 80470,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 335500,
        "netPrice": 100700,
        "costPrice": 61900
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 291700,
        "netPrice": 88000,
        "costPrice": 59950
      }
    ]
  },
  {
    "id": "REN-19-031",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 400A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 857870,
    "netPrice": 257361,
    "costPrice": 158470,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 659900,
        "netPrice": 198000,
        "costPrice": 121900
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 573800,
        "netPrice": 172200,
        "costPrice": 118050
      }
    ]
  },
  {
    "id": "REN-19-032",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 600A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 1249820,
    "netPrice": 374946,
    "costPrice": 229840,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "15%up R6.10.1",
        "effectiveDate": "2024-10-01",
        "listPrice": 961400,
        "netPrice": 288500,
        "costPrice": 176800
      },
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 836000,
        "netPrice": 250800,
        "costPrice": 171250
      }
    ]
  },
  {
    "id": "REN-19-033",
    "categoryId": "cat-18",
    "category": "ELB(FP)",
    "name": "ELB 3P 800A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": null,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "改定前",
        "effectiveDate": null,
        "listPrice": 1093300,
        "netPrice": 328000,
        "costPrice": 226030
      }
    ]
  },
  {
    "id": "REN-20-005",
    "categoryId": "cat-19",
    "category": "ｽﾘﾑ・KC",
    "name": "MCB 2P1E 15/20/30A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 5447,
    "netPrice": 1634,
    "costPrice": 910,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "ｽﾘﾑ・KC",
        "effectiveDate": null,
        "listPrice": 4190,
        "netPrice": 1300,
        "costPrice": 700
      }
    ]
  },
  {
    "id": "REN-20-006",
    "categoryId": "cat-19",
    "category": "ｽﾘﾑ・KC",
    "name": "MCB 2P2E 15/20/30A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 7774,
    "netPrice": 2332,
    "costPrice": 1300,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "ｽﾘﾑ・KC",
        "effectiveDate": null,
        "listPrice": 5980,
        "netPrice": 1800,
        "costPrice": 1000
      }
    ]
  },
  {
    "id": "REN-20-008",
    "categoryId": "cat-19",
    "category": "ｽﾘﾑ・KC",
    "name": "ELB 2P1E 15/20/30A 15/30mA",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 14950,
    "netPrice": 4485,
    "costPrice": 3354,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "ｽﾘﾑ・KC",
        "effectiveDate": null,
        "listPrice": 11500,
        "netPrice": 3500,
        "costPrice": 2580
      }
    ]
  },
  {
    "id": "REN-20-009",
    "categoryId": "cat-19",
    "category": "ｽﾘﾑ・KC",
    "name": "ELB 2P2E 15/20/30A 15/30mA",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 21580,
    "netPrice": 6474,
    "costPrice": 4797,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "ｽﾘﾑ・KC",
        "effectiveDate": null,
        "listPrice": 16600,
        "netPrice": 5000,
        "costPrice": 3690
      }
    ]
  },
  {
    "id": "REN-20-013",
    "categoryId": "cat-19",
    "category": "ｽﾘﾑ・KC",
    "name": "NF30-KC 3P 3～30A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 14950,
    "netPrice": 4485,
    "costPrice": 2561,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "ｽﾘﾑ・KC",
        "effectiveDate": null,
        "listPrice": 11500,
        "netPrice": 3500,
        "costPrice": 1970
      }
    ]
  },
  {
    "id": "REN-20-014",
    "categoryId": "cat-19",
    "category": "ｽﾘﾑ・KC",
    "name": "NF50-KC 3P 5～50A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 20800,
    "netPrice": 6240,
    "costPrice": 3497,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "ｽﾘﾑ・KC",
        "effectiveDate": null,
        "listPrice": 16000,
        "netPrice": 4800,
        "costPrice": 2690
      }
    ]
  },
  {
    "id": "REN-20-015",
    "categoryId": "cat-19",
    "category": "ｽﾘﾑ・KC",
    "name": "NF100-KC 3P 60～100A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 45760,
    "netPrice": 13728,
    "costPrice": 7852,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "ｽﾘﾑ・KC",
        "effectiveDate": null,
        "listPrice": 35200,
        "netPrice": 10600,
        "costPrice": 6040
      }
    ]
  },
  {
    "id": "REN-20-017",
    "categoryId": "cat-19",
    "category": "ｽﾘﾑ・KC",
    "name": "NV30-KC 3P 5～30A 15/30/100mA",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 28990,
    "netPrice": 8697,
    "costPrice": 6370,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "ｽﾘﾑ・KC",
        "effectiveDate": null,
        "listPrice": 22300,
        "netPrice": 6700,
        "costPrice": 4900
      }
    ]
  },
  {
    "id": "REN-20-018",
    "categoryId": "cat-19",
    "category": "ｽﾘﾑ・KC",
    "name": "NV50-KC 3P 5～50A 15/30/100mA",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 29770,
    "netPrice": 8931,
    "costPrice": 6682,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "ｽﾘﾑ・KC",
        "effectiveDate": null,
        "listPrice": 22900,
        "netPrice": 6900,
        "costPrice": 5140
      }
    ]
  },
  {
    "id": "REN-20-019",
    "categoryId": "cat-19",
    "category": "ｽﾘﾑ・KC",
    "name": "NV100-KC 3P 60～100A 30/100mA",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 58630,
    "netPrice": 17589,
    "costPrice": 13091,
    "priceLabel": "30%up R8.7.1",
    "effectiveDate": "2026-07-01",
    "history": [
      {
        "label": "ｽﾘﾑ・KC",
        "effectiveDate": null,
        "listPrice": 45100,
        "netPrice": 13600,
        "costPrice": 10070
      }
    ]
  },
  {
    "id": "REN-21-005",
    "categoryId": "cat-20",
    "category": "特殊Br・DCS",
    "name": "NF400-ZCW 3P 400A 埋込形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 183700,
    "netPrice": 88000,
    "costPrice": 56100,
    "priceLabel": "15%up",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-21-006",
    "categoryId": "cat-20",
    "category": "特殊Br・DCS",
    "name": "NF800-ZEW 3P 800A 埋込形",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 674300,
    "netPrice": 284000,
    "costPrice": 181150,
    "priceLabel": "15%up",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-21-012",
    "categoryId": "cat-20",
    "category": "特殊Br・DCS",
    "name": "DCS 3P 200A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 50500,
    "netPrice": 25730,
    "costPrice": 25730,
    "priceLabel": "15%up",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-21-013",
    "categoryId": "cat-20",
    "category": "特殊Br・DCS",
    "name": "DCS 3P 300A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 109000,
    "netPrice": 55580,
    "costPrice": 55580,
    "priceLabel": "15%up",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-21-014",
    "categoryId": "cat-20",
    "category": "特殊Br・DCS",
    "name": "DCS 3P 400A",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 159000,
    "netPrice": 81000,
    "costPrice": 81000,
    "priceLabel": "15%up",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-22-004",
    "categoryId": "cat-21",
    "category": "Panasonic",
    "name": "多回路エネルギーモニタ 本体ユニット",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 78400,
    "netPrice": 48000,
    "costPrice": 47970,
    "priceLabel": "黄色の箇所 2026/2/6 定価・仕入れとも約1.1倍",
    "effectiveDate": "2026-02-06",
    "history": []
  },
  {
    "id": "REN-22-005",
    "categoryId": "cat-21",
    "category": "Panasonic",
    "name": "多回路エネルギーモニタ 本体ユニット",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 78400,
    "netPrice": 48000,
    "costPrice": 47970,
    "priceLabel": "黄色の箇所 2026/2/6 定価・仕入れとも約1.1倍",
    "effectiveDate": "2026-02-06",
    "history": []
  },
  {
    "id": "REN-22-006",
    "categoryId": "cat-21",
    "category": "Panasonic",
    "name": "多回路エネルギモニタ 増設ユニット",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 31500,
    "netPrice": 20000,
    "costPrice": 19260,
    "priceLabel": "黄色の箇所 2026/2/6 定価・仕入れとも約1.1倍",
    "effectiveDate": "2026-02-06",
    "history": []
  },
  {
    "id": "REN-22-007",
    "categoryId": "cat-21",
    "category": "Panasonic",
    "name": "多回路エネルギーモニタ異種系統ユニット",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 56000,
    "netPrice": 34200,
    "costPrice": 34200,
    "priceLabel": "黄色の箇所 2026/2/6 定価・仕入れとも約1.1倍",
    "effectiveDate": "2026-02-06",
    "history": []
  },
  {
    "id": "REN-22-008",
    "categoryId": "cat-21",
    "category": "Panasonic",
    "name": "多回路エネルギーモニタ．表示設定ＭＢ１",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 38500,
    "netPrice": 23490,
    "costPrice": 23490,
    "priceLabel": "黄色の箇所 2026/2/6 定価・仕入れとも約1.1倍",
    "effectiveDate": "2026-02-06",
    "history": []
  },
  {
    "id": "REN-22-009",
    "categoryId": "cat-21",
    "category": "Panasonic",
    "name": "エネルギーモニタ ＬＡＮ対応ユニット",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 88000,
    "netPrice": 53820,
    "costPrice": 53820,
    "priceLabel": "黄色の箇所 2026/2/6 定価・仕入れとも約1.1倍",
    "effectiveDate": "2026-02-06",
    "history": []
  },
  {
    "id": "REN-22-010",
    "categoryId": "cat-21",
    "category": "Panasonic",
    "name": "エネルギーモニタ ＣＴ ６００Ａ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 24200,
    "netPrice": 14805,
    "costPrice": 14805,
    "priceLabel": "黄色の箇所 2026/2/6 定価・仕入れとも約1.1倍",
    "effectiveDate": "2026-02-06",
    "history": []
  },
  {
    "id": "REN-22-011",
    "categoryId": "cat-21",
    "category": "Panasonic",
    "name": "エネルギーモニタ ＣＴ ２５０Ａ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 12100,
    "netPrice": 7389,
    "costPrice": 7389,
    "priceLabel": "黄色の箇所 2026/2/6 定価・仕入れとも約1.1倍",
    "effectiveDate": "2026-02-06",
    "history": []
  },
  {
    "id": "REN-22-012",
    "categoryId": "cat-21",
    "category": "Panasonic",
    "name": "エネルギーモニタ ＣＴ １００Ａ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 7280,
    "netPrice": 5000,
    "costPrice": 4460,
    "priceLabel": "黄色の箇所 2026/2/6 定価・仕入れとも約1.1倍",
    "effectiveDate": "2026-02-06",
    "history": []
  },
  {
    "id": "REN-22-013",
    "categoryId": "cat-21",
    "category": "Panasonic",
    "name": "エネルギーモニタ ＣＴ ５０Ａ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 6030,
    "netPrice": 4000,
    "costPrice": 3690,
    "priceLabel": "黄色の箇所 2026/2/6 定価・仕入れとも約1.1倍",
    "effectiveDate": "2026-02-06",
    "history": []
  },
  {
    "id": "REN-22-014",
    "categoryId": "cat-21",
    "category": "Panasonic",
    "name": "多回路エネルギーモニタＣＴケーブル１０ｍ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 6600,
    "netPrice": 4050,
    "costPrice": 4050,
    "priceLabel": "黄色の箇所 2026/2/6 定価・仕入れとも約1.1倍",
    "effectiveDate": "2026-02-06",
    "history": []
  },
  {
    "id": "REN-22-015",
    "categoryId": "cat-21",
    "category": "Panasonic",
    "name": "多回路エネルギーモニタ．ＣＴケーブル５ｍ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 4860,
    "netPrice": 3000,
    "costPrice": 2970,
    "priceLabel": "黄色の箇所 2026/2/6 定価・仕入れとも約1.1倍",
    "effectiveDate": "2026-02-06",
    "history": []
  },
  {
    "id": "REN-22-016",
    "categoryId": "cat-21",
    "category": "Panasonic",
    "name": "多回路エネルギーモニタ．ＣＴケーブル２ｍ",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 3300,
    "netPrice": 3000,
    "costPrice": 2020,
    "priceLabel": "黄色の箇所 2026/2/6 定価・仕入れとも約1.1倍",
    "effectiveDate": "2026-02-06",
    "history": []
  },
  {
    "id": "REN-22-018",
    "categoryId": "cat-21",
    "category": "Panasonic",
    "name": "Ｆ２ 伝送ユニット",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 71000,
    "netPrice": 45090,
    "costPrice": 45090,
    "priceLabel": "黄色の箇所 2026/2/6 定価・仕入れとも約1.1倍",
    "effectiveDate": "2026-02-06",
    "history": []
  },
  {
    "id": "REN-22-022",
    "categoryId": "cat-21",
    "category": "Panasonic",
    "name": "WR2301 リモコントランス",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": 6000,
    "costPrice": 5580,
    "priceLabel": "黄色の箇所 2026/2/6 定価・仕入れとも約1.1倍",
    "effectiveDate": "2026-02-06",
    "history": []
  },
  {
    "id": "REN-22-023",
    "categoryId": "cat-21",
    "category": "Panasonic",
    "name": "WR3400 リレー制御用T/U",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": 4000,
    "costPrice": 3400,
    "priceLabel": "黄色の箇所 2026/2/6 定価・仕入れとも約1.1倍",
    "effectiveDate": "2026-02-06",
    "history": []
  },
  {
    "id": "REN-22-024",
    "categoryId": "cat-21",
    "category": "Panasonic",
    "name": "WR6165 リモコンリレー(片切)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": 3000,
    "costPrice": 2187,
    "priceLabel": "黄色の箇所 2026/2/6 定価・仕入れとも約1.1倍",
    "effectiveDate": "2026-02-06",
    "history": []
  },
  {
    "id": "REN-22-025",
    "categoryId": "cat-21",
    "category": "Panasonic",
    "name": "WR6166 リモコンリレー(両切)",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": 5000,
    "costPrice": 4050,
    "priceLabel": "黄色の箇所 2026/2/6 定価・仕入れとも約1.1倍",
    "effectiveDate": "2026-02-06",
    "history": []
  },
  {
    "id": "REN-22-026",
    "categoryId": "cat-21",
    "category": "Panasonic",
    "name": "WRT2050 伝送ユニット",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": 46000,
    "costPrice": 45090,
    "priceLabel": "黄色の箇所 2026/2/6 定価・仕入れとも約1.1倍",
    "effectiveDate": "2026-02-06",
    "history": []
  },
  {
    "id": "REN-22-027",
    "categoryId": "cat-21",
    "category": "Panasonic",
    "name": "WRT3224K 接点入力T/U",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": 10000,
    "costPrice": 9270,
    "priceLabel": "黄色の箇所 2026/2/6 定価・仕入れとも約1.1倍",
    "effectiveDate": "2026-02-06",
    "history": []
  },
  {
    "id": "REN-22-028",
    "categoryId": "cat-21",
    "category": "Panasonic",
    "name": "WRT3540K 年間プログラムタイマー",
    "spec": "",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": 45000,
    "costPrice": 44280,
    "priceLabel": "黄色の箇所 2026/2/6 定価・仕入れとも約1.1倍",
    "effectiveDate": "2026-02-06",
    "history": []
  },
  {
    "id": "REN-MISC-001",
    "categoryId": "cat-misc",
    "category": "その他",
    "name": "NF800-SEW 3P 800A 埋込型FP 電動式NFM",
    "spec": "納期目安 1.5～2か月。操作電圧要確認",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": 198880,
    "priceLabel": "個別記録",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-MISC-002",
    "categoryId": "cat-misc",
    "category": "その他",
    "name": "PTUV10160 引出形VT",
    "spec": "別紙仕様。納期目安 4～5か月",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": 354600,
    "priceLabel": "個別記録",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-MISC-003",
    "categoryId": "cat-misc",
    "category": "その他",
    "name": "NF1600-SEW 3P 1500A",
    "spec": "納期目安 1～1.5か月",
    "defaultQty": 1,
    "listPrice": null,
    "netPrice": null,
    "costPrice": 289610,
    "priceLabel": "個別記録",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-MISC-004",
    "categoryId": "cat-misc",
    "category": "その他",
    "name": "FR-D820-2.2K-100",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 118400,
    "netPrice": null,
    "costPrice": 24860,
    "priceLabel": "個別記録",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-MISC-005",
    "categoryId": "cat-misc",
    "category": "その他",
    "name": "E-30S5",
    "spec": "大垣市場",
    "defaultQty": 1,
    "listPrice": 52900,
    "netPrice": null,
    "costPrice": 20680,
    "priceLabel": "個別記録",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-MISC-006",
    "categoryId": "cat-misc",
    "category": "その他",
    "name": "EX-30A2",
    "spec": "日本昭和音楽村",
    "defaultQty": 1,
    "listPrice": 45500,
    "netPrice": null,
    "costPrice": 17790,
    "priceLabel": "個別記録",
    "effectiveDate": null,
    "history": []
  },
  {
    "id": "REN-MISC-007",
    "categoryId": "cat-misc",
    "category": "その他",
    "name": "三菱 有圧換気扇 AC100V 25cm EF-25UAS2",
    "spec": "",
    "defaultQty": 1,
    "listPrice": 45400,
    "netPrice": null,
    "costPrice": 13890,
    "priceLabel": "個別記録",
    "effectiveDate": null,
    "history": []
  }
];

const RENOVATION_MATERIAL_GROUPS = [
  "電線(KIP)",
  "電線(IV)",
  "電線(HIV)",
  "電線(MLFC)",
  "盤内材料",
  "ブランク",
  "銅バー",
  "鋼材・板材",
  "その他"
];

const RENOVATION_MATERIALS = [
  {
    "id": "MAT-IV-8",
    "group": "電線",
    "name": "IV電線",
    "spec": "8sq",
    "unit": "m",
    "unitPrice": 200
  },
  {
    "id": "MAT-IV-14",
    "group": "電線",
    "name": "IV電線",
    "spec": "14sq",
    "unit": "m",
    "unitPrice": 300
  },
  {
    "id": "MAT-IV-22",
    "group": "電線",
    "name": "IV電線",
    "spec": "22sq",
    "unit": "m",
    "unitPrice": 400
  },
  {
    "id": "MAT-IV-38",
    "group": "電線",
    "name": "IV電線",
    "spec": "38sq",
    "unit": "m",
    "unitPrice": 700
  },
  {
    "id": "MAT-IV-60",
    "group": "電線",
    "name": "IV電線",
    "spec": "60sq",
    "unit": "m",
    "unitPrice": 1100
  },
  {
    "id": "MAT-IV-100",
    "group": "電線",
    "name": "IV電線",
    "spec": "100sq",
    "unit": "m",
    "unitPrice": 1900
  },
  {
    "id": "MAT-IV-150",
    "group": "電線",
    "name": "IV電線",
    "spec": "150sq",
    "unit": "m",
    "unitPrice": 2900
  },
  {
    "id": "MAT-IV-200",
    "group": "電線",
    "name": "IV電線",
    "spec": "200sq",
    "unit": "m",
    "unitPrice": 3600
  },
  {
    "id": "MAT-IV-250",
    "group": "電線",
    "name": "IV電線",
    "spec": "250sq",
    "unit": "m",
    "unitPrice": 4700
  },
  {
    "id": "MAT-IV-325",
    "group": "電線",
    "name": "IV電線",
    "spec": "325sq",
    "unit": "m",
    "unitPrice": 6000
  },
  {
    "id": "MAT-HIV-8",
    "group": "電線",
    "name": "HIV電線",
    "spec": "8sq",
    "unit": "m",
    "unitPrice": 300
  },
  {
    "id": "MAT-HIV-14",
    "group": "電線",
    "name": "HIV電線",
    "spec": "14sq",
    "unit": "m",
    "unitPrice": 500
  },
  {
    "id": "MAT-HIV-22",
    "group": "電線",
    "name": "HIV電線",
    "spec": "22sq",
    "unit": "m",
    "unitPrice": 800
  },
  {
    "id": "MAT-HIV-38",
    "group": "電線",
    "name": "HIV電線",
    "spec": "38sq",
    "unit": "m",
    "unitPrice": 1300
  },
  {
    "id": "MAT-HIV-60",
    "group": "電線",
    "name": "HIV電線",
    "spec": "60sq",
    "unit": "m",
    "unitPrice": 2100
  },
  {
    "id": "MAT-HIV-100",
    "group": "電線",
    "name": "HIV電線",
    "spec": "100sq",
    "unit": "m",
    "unitPrice": 3400
  },
  {
    "id": "MAT-HIV-150",
    "group": "電線",
    "name": "HIV電線",
    "spec": "150sq",
    "unit": "m",
    "unitPrice": 4000
  },
  {
    "id": "MAT-HIV-200",
    "group": "電線",
    "name": "HIV電線",
    "spec": "200sq",
    "unit": "m",
    "unitPrice": 5000
  },
  {
    "id": "MAT-HIV-250",
    "group": "電線",
    "name": "HIV電線",
    "spec": "250sq",
    "unit": "m",
    "unitPrice": 6000
  },
  {
    "id": "MAT-HIV-325",
    "group": "電線",
    "name": "HIV電線",
    "spec": "325sq",
    "unit": "m",
    "unitPrice": 8000
  },
  {
    "id": "MAT-MLFC-8",
    "group": "電線",
    "name": "MLFC電線",
    "spec": "8sq",
    "unit": "m",
    "unitPrice": 350
  },
  {
    "id": "MAT-MLFC-14",
    "group": "電線",
    "name": "MLFC電線",
    "spec": "14sq",
    "unit": "m",
    "unitPrice": 600
  },
  {
    "id": "MAT-MLFC-22",
    "group": "電線",
    "name": "MLFC電線",
    "spec": "22sq",
    "unit": "m",
    "unitPrice": 900
  },
  {
    "id": "MAT-MLFC-38",
    "group": "電線",
    "name": "MLFC電線",
    "spec": "38sq",
    "unit": "m",
    "unitPrice": 1500
  },
  {
    "id": "MAT-MLFC-60",
    "group": "電線",
    "name": "MLFC電線",
    "spec": "60sq",
    "unit": "m",
    "unitPrice": 2500
  },
  {
    "id": "MAT-MLFC-100",
    "group": "電線",
    "name": "MLFC電線",
    "spec": "100sq",
    "unit": "m",
    "unitPrice": 4000
  },
  {
    "id": "MAT-MLFC-150",
    "group": "電線",
    "name": "MLFC電線",
    "spec": "150sq",
    "unit": "m",
    "unitPrice": 5500
  },
  {
    "id": "MAT-MLFC-200",
    "group": "電線",
    "name": "MLFC電線",
    "spec": "200sq",
    "unit": "m",
    "unitPrice": 7000
  },
  {
    "id": "MAT-MLFC-250",
    "group": "電線",
    "name": "MLFC電線",
    "spec": "250sq",
    "unit": "m",
    "unitPrice": 9000
  },
  {
    "id": "MAT-MLFC-325",
    "group": "電線",
    "name": "MLFC電線",
    "spec": "325sq",
    "unit": "m",
    "unitPrice": 12000
  },
  {
    "id": "MAT-KIP-8",
    "group": "電線",
    "name": "KIP電線",
    "spec": "8sq",
    "unit": "m",
    "unitPrice": 500
  },
  {
    "id": "MAT-KIP-14",
    "group": "電線",
    "name": "KIP電線",
    "spec": "14sq",
    "unit": "m",
    "unitPrice": 800
  },
  {
    "id": "MAT-KIP-22",
    "group": "電線",
    "name": "KIP電線",
    "spec": "22sq",
    "unit": "m",
    "unitPrice": 1500
  },
  {
    "id": "MAT-KIP-38",
    "group": "電線",
    "name": "KIP電線",
    "spec": "38sq",
    "unit": "m",
    "unitPrice": 2000
  },
  {
    "id": "MAT-KIP-60",
    "group": "電線",
    "name": "KIP電線",
    "spec": "60sq",
    "unit": "m",
    "unitPrice": 3000
  },
  {
    "id": "MAT-KIP-100",
    "group": "電線",
    "name": "KIP電線",
    "spec": "100sq",
    "unit": "m",
    "unitPrice": 5000
  },
  {
    "id": "MAT-KIP-150",
    "group": "電線",
    "name": "KIP電線",
    "spec": "150sq",
    "unit": "m",
    "unitPrice": 7000
  },
  {
    "id": "MAT-KIP-200",
    "group": "電線",
    "name": "KIP電線",
    "spec": "200sq",
    "unit": "m",
    "unitPrice": 10000
  },
  {
    "id": "MAT-KIP-250",
    "group": "電線",
    "name": "KIP電線",
    "spec": "250sq",
    "unit": "m",
    "unitPrice": 12500
  },
  {
    "id": "MAT-KIP-325",
    "group": "電線",
    "name": "KIP電線",
    "spec": "325sq",
    "unit": "m",
    "unitPrice": null
  },
  {
    "id": "MAT-PART-1",
    "group": "盤内材料",
    "name": "架台",
    "spec": "",
    "unit": "個",
    "unitPrice": 10000
  },
  {
    "id": "MAT-PART-2",
    "group": "盤内材料",
    "name": "長孔",
    "spec": "",
    "unit": "個",
    "unitPrice": 3000
  },
  {
    "id": "MAT-PART-3",
    "group": "盤内材料",
    "name": "碍子",
    "spec": "",
    "unit": "個",
    "unitPrice": 3000
  },
  {
    "id": "MAT-PART-4",
    "group": "盤内材料",
    "name": "クリート",
    "spec": "大",
    "unit": "個",
    "unitPrice": 2000
  },
  {
    "id": "MAT-PART-5",
    "group": "盤内材料",
    "name": "クリート",
    "spec": "小",
    "unit": "個",
    "unitPrice": 1000
  },
  {
    "id": "MAT-BLANK-1",
    "group": "ブランク",
    "name": "ブランク",
    "spec": "VCB",
    "unit": "箇所",
    "unitPrice": 10000
  },
  {
    "id": "MAT-BLANK-2",
    "group": "ブランク",
    "name": "ブランク",
    "spec": "OCR",
    "unit": "箇所",
    "unitPrice": 5000
  },
  {
    "id": "MAT-BLANK-3",
    "group": "ブランク",
    "name": "ブランク",
    "spec": "M",
    "unit": "箇所",
    "unitPrice": 2000
  },
  {
    "id": "MAT-H-3",
    "group": "鋼材・板材",
    "name": "H鋼",
    "spec": "200×200",
    "unit": "m",
    "unitPrice": 44910
  },
  {
    "id": "MAT-H-4",
    "group": "鋼材・板材",
    "name": "H鋼",
    "spec": "150×150",
    "unit": "m",
    "unitPrice": 27990
  },
  {
    "id": "MAT-H-5",
    "group": "鋼材・板材",
    "name": "H鋼",
    "spec": "300×300",
    "unit": "m",
    "unitPrice": 83700
  },
  {
    "id": "MAT-CU-9",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "2×12",
    "unit": "m",
    "unitPrice": 2220
  },
  {
    "id": "MAT-CU-10",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "2×20",
    "unit": "m",
    "unitPrice": 3560
  },
  {
    "id": "MAT-CU-11",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "3×12",
    "unit": "m",
    "unitPrice": 2800
  },
  {
    "id": "MAT-CU-12",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "3×15",
    "unit": "m",
    "unitPrice": 3440
  },
  {
    "id": "MAT-CU-13",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "3×18",
    "unit": "m",
    "unitPrice": 4080
  },
  {
    "id": "MAT-CU-14",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "3×20",
    "unit": "m",
    "unitPrice": 4540
  },
  {
    "id": "MAT-CU-15",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "3×25",
    "unit": "m",
    "unitPrice": 5640
  },
  {
    "id": "MAT-CU-16",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "3×30",
    "unit": "m",
    "unitPrice": 6640
  },
  {
    "id": "MAT-CU-17",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "3×40",
    "unit": "m",
    "unitPrice": 8840
  },
  {
    "id": "MAT-CU-18",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "3×50",
    "unit": "m",
    "unitPrice": 10940
  },
  {
    "id": "MAT-CU-19",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "4×15",
    "unit": "m",
    "unitPrice": 4220
  },
  {
    "id": "MAT-CU-20",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "4×20",
    "unit": "m",
    "unitPrice": 5520
  },
  {
    "id": "MAT-CU-21",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "4×25",
    "unit": "m",
    "unitPrice": 6820
  },
  {
    "id": "MAT-CU-22",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "4×30",
    "unit": "m",
    "unitPrice": 8120
  },
  {
    "id": "MAT-CU-23",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "4×40",
    "unit": "m",
    "unitPrice": 10720
  },
  {
    "id": "MAT-CU-24",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "4×50",
    "unit": "m",
    "unitPrice": 13220
  },
  {
    "id": "MAT-CU-25",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "5×20",
    "unit": "m",
    "unitPrice": 6500
  },
  {
    "id": "MAT-CU-26",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "5×25",
    "unit": "m",
    "unitPrice": 8000
  },
  {
    "id": "MAT-CU-27",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "5×30",
    "unit": "m",
    "unitPrice": 9500
  },
  {
    "id": "MAT-CU-28",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "5×32",
    "unit": "m",
    "unitPrice": 10160
  },
  {
    "id": "MAT-CU-29",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "5×40",
    "unit": "m",
    "unitPrice": 12500
  },
  {
    "id": "MAT-CU-30",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "5×50",
    "unit": "m",
    "unitPrice": 15600
  },
  {
    "id": "MAT-CU-31",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "5×75",
    "unit": "m",
    "unitPrice": 23100
  },
  {
    "id": "MAT-CU-32",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "6×20",
    "unit": "m",
    "unitPrice": 7480
  },
  {
    "id": "MAT-CU-33",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "6×25",
    "unit": "m",
    "unitPrice": 9180
  },
  {
    "id": "MAT-CU-34",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "6×30",
    "unit": "m",
    "unitPrice": 10880
  },
  {
    "id": "MAT-CU-35",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "6×32",
    "unit": "m",
    "unitPrice": 11640
  },
  {
    "id": "MAT-CU-36",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "6×35",
    "unit": "m",
    "unitPrice": 12680
  },
  {
    "id": "MAT-CU-37",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "6×38",
    "unit": "m",
    "unitPrice": 13720
  },
  {
    "id": "MAT-CU-38",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "6×40",
    "unit": "m",
    "unitPrice": 14380
  },
  {
    "id": "MAT-CU-39",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "6×45",
    "unit": "m",
    "unitPrice": 16080
  },
  {
    "id": "MAT-CU-40",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "6×50",
    "unit": "m",
    "unitPrice": 17880
  },
  {
    "id": "MAT-CU-41",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "6×60",
    "unit": "m",
    "unitPrice": 21280
  },
  {
    "id": "MAT-CU-42",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "6×70",
    "unit": "m",
    "unitPrice": 24780
  },
  {
    "id": "MAT-CU-43",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "6×75",
    "unit": "m",
    "unitPrice": 26480
  },
  {
    "id": "MAT-CU-44",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "6×100",
    "unit": "m",
    "unitPrice": 35180
  },
  {
    "id": "MAT-CU-45",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "6×150",
    "unit": "m",
    "unitPrice": 52480
  },
  {
    "id": "MAT-CU-46",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "6×200",
    "unit": "m",
    "unitPrice": 69880
  },
  {
    "id": "MAT-CU-47",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "8×25",
    "unit": "m",
    "unitPrice": 11540
  },
  {
    "id": "MAT-CU-48",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "8×30",
    "unit": "m",
    "unitPrice": 13740
  },
  {
    "id": "MAT-CU-49",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "8×40",
    "unit": "m",
    "unitPrice": 18140
  },
  {
    "id": "MAT-CU-50",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "8×50",
    "unit": "m",
    "unitPrice": 22440
  },
  {
    "id": "MAT-CU-51",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "8×60",
    "unit": "m",
    "unitPrice": 26840
  },
  {
    "id": "MAT-CU-52",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "8×65",
    "unit": "m",
    "unitPrice": 29040
  },
  {
    "id": "MAT-CU-53",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "8×75",
    "unit": "m",
    "unitPrice": 33340
  },
  {
    "id": "MAT-CU-54",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "8×100",
    "unit": "m",
    "unitPrice": 44240
  },
  {
    "id": "MAT-CU-55",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "8×150",
    "unit": "m",
    "unitPrice": 66040
  },
  {
    "id": "MAT-CU-56",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "8×200",
    "unit": "m",
    "unitPrice": 87740
  },
  {
    "id": "MAT-CU-57",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "10×30",
    "unit": "m",
    "unitPrice": 16600
  },
  {
    "id": "MAT-CU-58",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "10×40",
    "unit": "m",
    "unitPrice": 21800
  },
  {
    "id": "MAT-CU-59",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "10×50",
    "unit": "m",
    "unitPrice": 27100
  },
  {
    "id": "MAT-CU-60",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "10×60",
    "unit": "m",
    "unitPrice": 32300
  },
  {
    "id": "MAT-CU-61",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "10×75",
    "unit": "m",
    "unitPrice": 40200
  },
  {
    "id": "MAT-CU-62",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "10×100",
    "unit": "m",
    "unitPrice": 53300
  },
  {
    "id": "MAT-CU-63",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "10×125",
    "unit": "m",
    "unitPrice": 66400
  },
  {
    "id": "MAT-CU-64",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "10×150",
    "unit": "m",
    "unitPrice": 79500
  },
  {
    "id": "MAT-CU-65",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "10×200",
    "unit": "m",
    "unitPrice": 105700
  },
  {
    "id": "MAT-CU-66",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "12×50",
    "unit": "m",
    "unitPrice": 31660
  },
  {
    "id": "MAT-CU-67",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "12×75",
    "unit": "m",
    "unitPrice": 46960
  },
  {
    "id": "MAT-CU-68",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "12×100",
    "unit": "m",
    "unitPrice": 62360
  },
  {
    "id": "MAT-CU-69",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "15×50",
    "unit": "m",
    "unitPrice": 38600
  },
  {
    "id": "MAT-CU-70",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "15×75",
    "unit": "m",
    "unitPrice": 57200
  },
  {
    "id": "MAT-CU-71",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "15×100",
    "unit": "m",
    "unitPrice": 75900
  },
  {
    "id": "MAT-CU-72",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "20×50",
    "unit": "m",
    "unitPrice": 50100
  },
  {
    "id": "MAT-CU-73",
    "group": "銅バー",
    "name": "銅バー",
    "spec": "20×100",
    "unit": "m",
    "unitPrice": 98500
  },
  {
    "id": "MAT-PLATE-1",
    "group": "鋼材・板材",
    "name": "鉄板",
    "spec": "1.6t",
    "unit": "㎡",
    "unitPrice": 14032
  },
  {
    "id": "MAT-PLATE-2",
    "group": "鋼材・板材",
    "name": "鉄板",
    "spec": "2.3t",
    "unit": "㎡",
    "unitPrice": 15792
  },
  {
    "id": "MAT-PLATE-3",
    "group": "鋼材・板材",
    "name": "鉄板",
    "spec": "3.2t",
    "unit": "㎡",
    "unitPrice": 18064
  },
  {
    "id": "MAT-PLATE-4",
    "group": "鋼材・板材",
    "name": "SUS板",
    "spec": "1.5t",
    "unit": "㎡",
    "unitPrice": 54000
  },
  {
    "id": "MAT-PLATE-5",
    "group": "鋼材・板材",
    "name": "SUS板",
    "spec": "2.0t",
    "unit": "㎡",
    "unitPrice": 35600
  },
  {
    "id": "MAT-PVC",
    "group": "鋼材・板材",
    "name": "塩ビ板",
    "spec": "寸法・穴あけ数で単価入力",
    "unit": "式",
    "unitPrice": null
  },
  {
    "id": "MAT-OTHER",
    "group": "その他",
    "name": "その他材料",
    "spec": "任意入力",
    "unit": "式",
    "unitPrice": null
  }
];

// 元Excel右側の「ブレーカー」欄と、添付資料の許容電流表を分けて保持する。
// breaker は実務上の選定目安、allowable は資料記載条件での参考許容電流。
const RENOVATION_WIRE_CAPACITY = {
  "IV電線": {
    8: { allowable: 61 }, 14: { allowable: 88 }, 22: { allowable: 115 }, 38: { allowable: 162 },
    60: { allowable: 217 }, 100: { allowable: 298 }, 150: { allowable: 395 }, 200: { allowable: 469 },
    250: { allowable: 556 }, 325: { allowable: 650 },
    condition: "600Vビニル絶縁電線・単線・周囲30℃（JIS C 4620:2023 解説表5）"
  },
  "HIV電線": {
    8: { breaker: 60, allowable: 66 }, 14: { breaker: 75, allowable: 95 }, 22: { breaker: 100, allowable: 124 },
    38: { breaker: 150, allowable: 175 }, 60: { breaker: 225, allowable: 234 }, 100: { breaker: 300, allowable: 322 },
    150: { breaker: 400, allowable: 426 }, 200: { breaker: 500, allowable: 506 }, 250: { breaker: 600, allowable: 600 },
    325: { breaker: 700, allowable: 702 },
    condition: "目安ブレーカーは元Excel。許容電流は公共建築工事標準仕様書（電気設備工事編）令和7年版 表1.1.6・基準周囲温度40℃"
  },
  "MLFC電線": {
    8: { breaker: 75, allowable: 93 }, 14: { breaker: 100, allowable: 134 }, 22: { breaker: 125, allowable: 175 },
    38: { breaker: 200, allowable: 247 }, 60: { breaker: 250, allowable: 332 }, 100: { breaker: 350, allowable: 455 },
    150: { breaker: 500, allowable: 604, parallel: "2本で800A目安" },
    200: { breaker: 600, allowable: 717, parallel: "2本で800A目安" },
    250: { breaker: 700, allowable: 850 }, 325: { breaker: 800, allowable: 994 },
    condition: "目安ブレーカーは元Excel。許容電流は導体許容温度110℃・周囲40℃・気中1条（メーカー資料）"
  },
  "KIP電線": {
    38: { breaker: 175 }, 60: { breaker: 235 }, 100: { breaker: 325 },
    condition: "元Excelに記載があるサイズのみ目安ブレーカーを表示"
  }
};

const RENOVATION_COPPER_CURRENT_DENSITY = [
  { maxCurrent: 400, density: 2.5 },
  { maxCurrent: 800, density: 2.0 },
  { maxCurrent: 1200, density: 1.7 },
  { maxCurrent: 2000, density: 1.5 },
  { maxCurrent: 2500, density: 1.3 }
];

const RENOVATION_REFERENCE_NOTES = {
  "general": [
    "試験費、処分費は含んでおりません。",
    "現地調査の結果、費用が変わる場合があります。",
    "重量運搬、クレーン車、重量物の撤去・据付は含んでおりません。"
  ],
  "patterns": {
    "transformer": [
      "重量運搬、クレーンおよび撤去品処分は別途です。"
    ],
    "matsuoka": [
      "トランスはトップランナーIIIで見積りしています。"
    ]
  },
  "categories": {
    "LBS": [
      "TB／TM／MG／RY／COS／パネル・箱は価格に含みません。"
    ],
    "PC・LA": [
      "PC-6-30（ヒューズ管付・無）、PC-6-50、PC-7-100は元表の『見積りはこっち』表記を優先して確認してください。"
    ],
    "その他": [
      "操作電圧・別紙仕様・納期の記載がある品目は、品目名と価格履歴を必ず確認してください。"
    ]
  },
  "materialGroups": {
    "電線(KIP)": [
      "KIP 325sqは元表で単価0のため未登録です。",
      "A表示は選定補助です。目安ブレーカーと資料記載条件の許容電流を区別し、周囲温度・布設条件・並列使用時は個別に確認してください。"
    ],
    "電線(IV)": [
      "A表示は選定補助です。資料記載条件の許容電流であり、周囲温度・布設条件は個別に確認してください。"
    ],
    "電線(HIV)": [
      "A表示は選定補助です。目安ブレーカーと資料記載条件の許容電流を区別し、周囲温度・布設条件は個別に確認してください。"
    ],
    "電線(MLFC)": [
      "A表示は選定補助です。目安ブレーカーとメーカー資料記載条件の許容電流を区別し、周囲温度・布設条件・並列使用時は個別に確認してください。"
    ],
    "盤内材料": [
      "元表の参考数量：受電盤は碍子3・クリート大21、コンデンサ盤は碍子3・クリート大12、トランスは碍子3・クリート大9。"
    ],
    "ブランク": [
      "元表の単価：VCB 10,000円、OCR 5,000円、M 2,000円。"
    ],
    "銅バー": [
      "銅2,500円/kg、メッキ200円/㎡、工賃係数2の元表計算を1m単価に反映しています。銅建値変更時は再確認してください。",
      "参考許容電流は公共建築工事標準仕様書（電気設備工事編）令和7年版 表1.1.5の電流密度から断面積で算出しています。材料の切取り・成形に対する+5%は含めていません。"
    ],
    "鋼材・板材": [
      "H鋼250円/kg・ドブ200円/kg、鉄板160円/kg・塗装5,000円/㎡など、元表の計算条件から登録しています。",
      "古河電工の可とう端子は定価なし。元表には銅建値と納期約1～1.5か月の記録があります。"
    ],
    "その他": [
      "元表にない材料は名称・数量・単価を確認して入力してください。"
    ]
  }
};
