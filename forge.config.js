module.exports = {
  packagerConfig: {
    name: "EterEnergy PDFtoExcel",
    asar: false,
    ignore: [
      "node_modules",
      "archivos/endesa/",
      "archivos/naturgy/",
      "archivos/nexus/",
      "archivos/repsol/",
      "archivos2/",
      "ExcelResumenFactura.xlsx",
      "ExcelResumenFactura2.xlsx",
    ],
    icon: 'assets/logo-etergy.ico',

  },
  rebuildConfig: {},
  makers: [
    {
      name: '@electron-forge/maker-squirrel',
      config: {
        icon: 'assets/logo-etergy.ico',
        setupIcon: 'assets/logo-etergy.ico',
      }

    }
  ],
  plugins: [],
};
