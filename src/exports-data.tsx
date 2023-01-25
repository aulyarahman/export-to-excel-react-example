import * as XLSX from "xlsx";
import * as ExcelJS from "exceljs";

const datas = [
  {
    stock_ndc_koli: 0.008771929824561403,
    stock_ndc_palet: 0,
    id: 1,
    odoo_code: "01911",
    kode_sl: "SA1",
    kode_item: "EB-IPB3M01-W5.5G",
    nama_item:
      "Wardah Instaperfect BROWFESSIONAL 3D Brow Mascara 01. BROWN 5.5 g",
    stock_ndc: 2,
    jenis_storage_ndc: "Trolly",
    streamline_id: 1,
    createdAt: "2023-01-12T07:08:14.906Z",
    updatedAt: "2023-01-12T07:08:14.906Z",
    deletedAt: null,
    mpq_kemas: {
      id: 1,
      pcs: 2456,
      pcs_per_troli: 228,
      troli: 10.771929824561404,
      troli_per_storage: 0,
      storage_ndc: 0,
      sku_id: 1,
      createdAt: "2023-01-12T07:08:14.912Z",
      updatedAt: "2023-01-12T07:08:14.912Z",
      deletedAt: null,
    },
    customer_order: {
      id: 1,
      pcs: 6,
      koli: 0.02631578947368421,
      palet: 0,
      sku_id: 1,
      createdAt: "2023-01-12T07:08:14.914Z",
      updatedAt: "2023-01-12T07:08:14.914Z",
      deletedAt: null,
    },
    leadtime: {
      id: 1,
      timbang: 1,
      premix: 2,
      olah_1: 2,
      olah_2: 2,
      kemas: 2,
      others: 2,
      sku_id: 1,
      createdAt: "2023-01-12T07:08:14.916Z",
      updatedAt: "2023-01-12T07:08:14.916Z",
      deletedAt: null,
    },
    critical_point: {
      id: 1,
      s_pcs: 1,
      s_koli: 0.0043859649122807015,
      s_ndc: null,
      rop_pcs: 2,
      rop_koli: 0.008771929824561403,
      rop_ndc: null,
      max_pcs: 2,
      max_koli: 0.008771929824561403,
      max_ndc: null,
      sku_id: 1,
      createdAt: "2023-01-12T07:08:14.918Z",
      updatedAt: "2023-01-12T07:08:14.918Z",
      deletedAt: null,
    },
  },
  {
    stock_ndc_koli: 0.008771929824561403,
    stock_ndc_palet: 0,
    id: 1,
    odoo_code: "01911",
    kode_sl: "SA1",
    kode_item: "EB-IPB3M01-W5.5G",
    nama_item:
      "Wardah Instaperfect BROWFESSIONAL 3D Brow Mascara 01. BROWN 5.5 g",
    stock_ndc: 2,
    jenis_storage_ndc: "Trolly",
    streamline_id: 1,
    createdAt: "2023-01-12T07:08:14.906Z",
    updatedAt: "2023-01-12T07:08:14.906Z",
    deletedAt: null,
    mpq_kemas: {
      id: 1,
      pcs: 2456,
      pcs_per_troli: 228,
      troli: 10.771929824561404,
      troli_per_storage: 0,
      storage_ndc: 0,
      sku_id: 1,
      createdAt: "2023-01-12T07:08:14.912Z",
      updatedAt: "2023-01-12T07:08:14.912Z",
      deletedAt: null,
    },
    customer_order: {
      id: 1,
      pcs: 6,
      koli: 0.02631578947368421,
      palet: 0,
      sku_id: 1,
      createdAt: "2023-01-12T07:08:14.914Z",
      updatedAt: "2023-01-12T07:08:14.914Z",
      deletedAt: null,
    },
    leadtime: {
      id: 1,
      timbang: 1,
      premix: 2,
      olah_1: 2,
      olah_2: 2,
      kemas: 2,
      others: 2,
      sku_id: 1,
      createdAt: "2023-01-12T07:08:14.916Z",
      updatedAt: "2023-01-12T07:08:14.916Z",
      deletedAt: null,
    },
    critical_point: {
      id: 1,
      s_pcs: 1,
      s_koli: 0.0043859649122807015,
      s_ndc: null,
      min_pcs: 2,
      min_koli: 0.008771929824561403,
      min_ndc: null,
      max_pcs: 2,
      max_koli: 0.008771929824561403,
      max_ndc: null,
      sku_id: 1,
      createdAt: "2023-01-12T07:08:14.918Z",
      updatedAt: "2023-01-12T07:08:14.918Z",
      deletedAt: null,
    },
  },
];

export const PagesExport = () => {
  const exportData = async () => {
    const newData: any[] = [];
    for await (const v of datas) {
      newData.push({
        "ODOO CODE": v.odoo_code,
        "KODE SL": v.kode_sl,
        "KODE ITEM": v.kode_item,
        "NAMA ITEM": v.nama_item,
        "DEDICATED STREAMLINE": v.jenis_storage_ndc,
        "NDC STOCK": {
          PCS: `${v.mpq_kemas.pcs}`,
          KOLI: `${v.mpq_kemas.pcs_per_troli}`,
          PALET: `${v.mpq_kemas.storage_ndc}`,
        },
        "CO MINGGUAN": {
          PCS: `${v.customer_order.pcs}`,
          "STORAGE NDC": `${v.customer_order.palet}`,
          KOLI: `${v.customer_order.koli}`,
        },
        "LEAD TIME (JAM)": {
          TIMBANG: v.leadtime.timbang,
          PREMIX: v.leadtime.premix,
          "OLAH 1": v.leadtime.olah_1,
          "OLAH 2": v.leadtime.olah_2,
          KEMAS: v.leadtime.kemas,
          "LAIN-LAIN": v.leadtime.others,
          TOTAL: v.leadtime.sku_id,
        },
        "_LEAD TIME (JAM)": {
          SAFETY: {
            PCS: v.critical_point.s_pcs,
            KOLI: v.critical_point.s_koli,
            "STRG NDC": v.critical_point.s_ndc,
          },
          ROP: {
            PCS: v.critical_point.rop_pcs,
            KOLI: v.critical_point.rop_koli,
            "STRG NDC": v.critical_point.rop_ndc,
          },
          MAXIMUM: {
            PCS: v.critical_point.max_pcs,
            KOLI: v.critical_point.max_koli,
            "STRG NDC": v.critical_point.max_ndc,
          },
        },
      });
    }

    const headers: string[] = [];
    Object.keys(newData[0]).forEach((v) => {
      headers.push(v);
    });

    headers.splice(headers.indexOf("NDC STOCK") + 1, 0, "", "");
    headers.splice(headers.indexOf("CO MINGGUAN") + 1, 0, "", "");
    headers.splice(
      headers.indexOf("LEAD TIME (JAM)") + 1,
      0,
      "",
      "",
      "",
      "",
      "",
      ""
    );

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Sheet1");
    worksheet.addRow(headers);
    worksheet.getCell("S1").value = "LEAD TIME (JAM)";

    const subHeader = {
      6: "PCS",
      7: "KOLI",
      8: "PALET",
      9: "KOLI",
      10: "PCS",
      11: "STORAGE NDC",
      12: "TIMBANG",
      13: "PREMIX",
      14: "OLAH 1",
      15: "OLAH 2",
      16: "KEMAS",
      17: "LAIN LAIN",
      18: "TOTAL",
      19: "SAFETY",
      20: "SAFETY",
      21: "SAFETY",
      22: "ROP",
      23: "ROP",
      24: "ROP",
      25: "MAXIMUM",
      26: "MAXIMUM",
      27: "MAXIMUM",
    };

    const subChildHeaders = {
      19: "PCS",
      20: "KOLI",
      21: "STRG NDC",
      22: "PCS",
      23: "KOLI",
      24: "STRG NDC",
      25: "PCS",
      26: "koli",
      27: "STRG NDC",
    };

    const cellsMerge = [
      "A1:A3",
      "B1:B3",
      "C1:C3",
      "D1:D3",
      "E1:E3",
      "F1:H1",
      "I1:K1",
      "F2:F3",
      "G2:G3",
      "H2:H3",
      "I2:I3",
      "J2:J3",
      "K2:K3",
      "L1:O1",
      "L2:L3",
      "M2:M3",
      "N2:N3",
      "O2:O3",
      "P2:P3",
      "Q2:Q3",
      "R2:R3",
      "S1:AA1",
    ];

    worksheet.addRow(Object.values(subHeader));

    worksheet.addRow(Object.values(subChildHeaders));

    cellsMerge.forEach((v) => worksheet.mergeCells(v));

    newData.forEach((item, k, arr) => {
      const arrKeys = arr.length + k + 2;
      worksheet.addRow(Object.values(item));

      Object.entries(item).forEach(([key, val]) => {
        const setKey =
          key === "NDC STOCK"
            ? "F,G,H"
            : key === "CO MINGGUAN"
            ? "I,J,K"
            : key === "LEAD TIME (JAM)"
            ? "L,M,N,O,P,Q,R"
            : "";

        if (typeof val === "object") {
          Object.entries(val as any).forEach(([key, vals], idx) => {
            const setKeys =
              key === "SAFETY"
                ? "S,T,U"
                : key === "ROP"
                ? "V,W,X"
                : key === "MAXIMUM"
                ? "Y,Z,AA"
                : "";

            if (typeof vals === "object") {
              Object.entries(val as any).forEach(([_, v]) => {
                Object.entries(v as any).forEach(([_, y], idx) => {
                  const s = setKeys.split(",")[idx];
                  worksheet.getCell(`${s}${arrKeys}`).value = y as any;
                });
              });
            } else {
              const s = setKey.split(",")[idx];
              worksheet.getCell(`${s}${arrKeys}`).value = vals as any;
            }
          });
        }
      });
    });

    for (let i = 1; i <= worksheet.columnCount; i++) {
      worksheet.getColumn(i).width = 20;
    }

    const selectedRows: any[] = [];
    worksheet.eachRow((row, rowNumber) => {
      if ([1, 2, 3].indexOf(rowNumber) !== -1) {
        selectedRows.push(row);
      }
    });

    selectedRows.forEach((row) => {
      row.eachCell((cell: any) => {
        cell.font = {
          bold: true,
        };
        cell.alignment = { horizontal: "center" };
      });
    });

    workbook.xlsx.writeBuffer().then((buffer) => {
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const link = document.createElement("a");
      link.href = window.URL.createObjectURL(blob);
      link.download = `sku_config_${new Date().toDateString()}.xlsx`;
      link.click();
    });
  };

  return <button onClick={exportData}>EXPORT</button>;
};
