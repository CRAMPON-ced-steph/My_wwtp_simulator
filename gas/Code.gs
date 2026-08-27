function doGet() {
  return HtmlService.createHtmlOutputFromFile('app')
    .setTitle('OCEAN — Simulateur filière eau')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
