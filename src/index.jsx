

import { Dropbox } from 'dropbox';
import ForgeUI, { render, ProjectPage, Fragment, Text, useEffect, useState, Button, window } from '@forge/ui';
import api, { route } from '@forge/api';
import XLSX from 'xlsx';
import { fetch } from '@forge/api';

 const dropboxAccessToken = 'sl.Beik6EXLYOvoRUXoX6t-saM_mpcMeFviy2bkqHYG1e7Nrj4AVdGdtwnYPIgLUIRG6eI4wmQO8w4Tu-H54MiCMkGLL8BCXTkjw-EhWxxCOdnm4rw35Jq2ywdgy16Ao2tj5obCJTQR01H5';
  const dropboxClient = new Dropbox({ accessToken: dropboxAccessToken });

async function fetchExcelFile() {
  console.log(`In the fetchExcelFile`);


  try {
    // Download the original Excel file from Dropbox
//     const filePath = '/MOD-JDI.xlsx';
    const filePath = '/MOD-JDI.xlsx';

   const dropboxApiUrl = 'https://content.dropboxapi.com/2/files/download';

      const headers = {
        'Authorization': `Bearer ${dropboxAccessToken}`,
        'Content-Type': 'application/octet-stream',
        'Dropbox-API-Arg': JSON.stringify({ path: filePath })
      };

  const response = await fetch(dropboxApiUrl, {
    method: 'POST',
    headers: headers
  });

  if (!response.ok) {
    console.error('Failed to download file from Dropbox');
    return null;
  }
  console.error('In fetchExcelFile after response.ok');

    const data = await response.arrayBuffer();
  console.error('In fetchExcelFile data  ' + data);
    // Read Excel file data
    const workbook = XLSX.read(data, { type: 'buffer' });
console.error('In fetchExcelFile workbook  ' + data);
//     // Assuming you want to work with the first sheet
//     const sheetName = workbook.SheetNames[0];
//     const sheet = workbook.Sheets[sheetName];
//
//     // Manipulate the sheet with your data
//     const yourData = 'Sample Data';
//     const cellAddress = 'A1'; // Replace with the desired cell address
//     sheet[cellAddress] = { v: yourData, t: 's' }; // 's' means 'string'
//  console.log(`In the fetchExcelFile before updatedData`);
    // Write the updated data back to the workbook
    const updatedData = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });


 console.log(`In the fetchExcelFile before uploadResponse`);
    // Upload the updated Excel file back to Dropbox
//     const uploadResponse = await dropboxClient.filesUpload({
//       path: '/updated-excel-file1.xlsx',
//       contents: updatedData,
//     });

 const filePath2 = '/test.xlsx';
    const headers2 = {
        'Authorization': `Bearer ${dropboxAccessToken}`,
        'Content-Type': 'application/octet-stream',
        'Dropbox-API-Arg': JSON.stringify({ path: filePath2 }),
         mode: 'overwrite'
      };
   const uploadResponse = await fetch('https://content.dropboxapi.com/2/files/upload', {
                                               method: 'post',
                                               headers: headers2,
                                               body: updatedData
                                             });


 console.log(`In the fetchExcelFile after uploadResponse`);
    return uploadResponse;
  } catch (error) {
    console.error('Error fetching Excel file:', error.error);
  }
}



// In the render method of the App component
function App() {
  const [isLoading, setIsLoading] = useState(false);
  const [shareableLink, setShareableLink] = useState(null);

  const handleButtonClick = async () => {
    setIsLoading(true);

    // Upload the updated Excel file to Dropbox and get its metadata
    const updatedExcelFileMetadata = await fetchExcelFile();
console.error('From app after fetchExcelFile:');
    // Fetch the shareable link of the updated Excel file
    const shareableLinkMetadata = await dropboxClient.sharingCreateSharedLinkWithSettings({
      path: updatedExcelFileMetadata.path_lower,
    });

    // Show the shareable link to the user
    setShareableLink(shareableLinkMetadata.url);

    setIsLoading(false);
  };

  return (
    <Fragment>
      <Text>Click the button below to generate the updated Excel file:</Text>
      <Button text="Generate" onClick={handleButtonClick} isLoading={isLoading} />
      {shareableLink && <Text>Click <a href={shareableLink}>here</a> to download the updated Excel file.</Text>}
    </Fragment>
  );

}

export const run = render(
    <ProjectPage>
        <App />
    </ProjectPage>
);
