module.exports = {
  packagerConfig: {
    name: "EterEnergy PDFtoExcel",
    asar: false,
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
