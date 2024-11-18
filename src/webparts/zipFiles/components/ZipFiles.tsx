/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-floating-promises */

import * as React from 'react';

import { Button } from '@fluentui/react-components';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { useZipFiles } from '../hooks/useZipFiles';

export const MyComponent = ({ context }: { context: WebPartContext }) => {
  const { createAndDownloadZip, ProgressDialog } = useZipFiles({ context });

  const handleDownload = () => {
    (async () => {
      const folderUrl = "/sites/ThePerspective/Shared Documents/DocumentSetApp"; // Example URL
      // const folderUrl =
      "https://graph.microsoft.com/v1.0/sites/spteck.sharepoint.com,1ee0a722-200d-4466-b31d-a0b6b5ae39c9,b06503c7-8710-4c0a-9b66-f2c36dc94a1b/drives/b!IqfgHg0gZkSzHaC2ta45yccDZbAQhwpMm2byw23JShtfdaSzvhi_T6au4LwUrIV-/root:/DocumentSetApp"; // Example URL
      await createAndDownloadZip(folderUrl, "my-files.zip");
    })();
  };

  return (
    <div>
      <Button onClick={handleDownload}>Download ZIP</Button>
      {ProgressDialog}
    </div>
  );
};
