


// CloudConvertAPI
function test_convertPdfToPng(){
  let file = DriveApp.getFileById("170Xfk33xfvLBheuu0ljdxCMgdH60_kmD")
  let fileUrl = "https://appsheet.al-pa-ca.com/wp-content/uploads/2025/01/meisi_sample1.pdf";//file.getDownloadUrl();
  let fileId = "test";
  const files = convertPdfToPng(fileUrl, fileId);
  console.log(files);
}
