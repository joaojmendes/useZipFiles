/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-floating-promises */

import * as React from 'react';

import {
  Body1,
  Button,
  Subtitle1,
} from '@fluentui/react-components';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { useZipFiles } from '../hooks/useZipFiles';

export const MyComponent = ({ context }: { context: WebPartContext }) => {
  const { createAndDownloadZip, ProgressDialog } = useZipFiles({ context });

  const handleDownload = () => {
    (async () => {
      const folderUrl = "/sites/ThePerspective/Shared Documents/DocumentSetApp"; // Example URL
      // const folderUrl =
      "https://graph.microsoft.com/v1.0/sites/spteck.sharepoint.com,1ee0a722-200d-4466-b31d-a0b6b5ae39c9,b06503c7-8710-4c0a-9b66-f2c36dc98888/drives/b!IqfgHg0gZkSzHaC2ta45yccDZbAQhwpMm2byw23JShtfdaSzvhi_T6au4Lw-/root:/DocumentSetApp"; // Example URL
      await createAndDownloadZip(folderUrl, "my-files.zip");
    })();
  };

  return (
    <div >
      <div style={{ display: "flex", flexDirection: "column", gap: 2 , paddingBottom: 10}}>
        <Subtitle1> Download ZIP file</Subtitle1>
        <Body1> Click the button below to download a ZIP file of the files in the folder,</Body1>{" "}
        <Body1 style={{ fontWeight: "bold" }}> /sites/ThePerspective/Shared Documents/DocumentSetApp </Body1>
      </div>
      <Button onClick={handleDownload}>Download ZIP</Button>
      {ProgressDialog}
    </div>
  );
};
