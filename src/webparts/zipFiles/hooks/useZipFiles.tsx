/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-use-before-define */
/* eslint-disable @typescript-eslint/no-floating-promises */
import * as React from 'react';

import { saveAs } from 'file-saver';
import * as JSZip from 'jszip';

import { css } from '@emotion/css';
import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  DialogTrigger,
  Field,
  FluentProvider,
  IdPrefixProvider,
  ProgressBar,
  Spinner,
  teamsDarkTheme,
  teamsHighContrastTheme,
  teamsLightTheme,
  Text,
  Theme,
  tokens,
  webDarkTheme,
  webLightTheme,
} from '@fluentui/react-components';
import {
  List,
  ListItem,
} from '@fluentui/react-list-preview';
import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

interface IFile {
  url: string;
  filename: string;
  folderName?: string;
}

interface UseZipFilesProps {
  context: WebPartContext;
  theme?: "webLight" | "webDark" | "teamsLight" | "teamsDark" | "teamsHighContrast";
}

const MAX_CONCURRENT_DOWNLOADS = 5; // Limit for parallel downloads
const MAX_RETRIES = 3; // Number of retry attempts for failed downloads

interface ErrorListProps {
  errors: string[];
}


const useStyles = () => {
  return {
    container: css`
      display: flex;
      flex-direction: column;
      gap: 10px;
    `,
    errorLogContainer: css`
      max-height: 150px;
      overflow-y: auto;
    `,
    errorText: css`
      color: ${tokens.colorStatusDangerForeground1};
    `,
    sucessText: css`
      color:  ${tokens.colorStatusSuccessBackground1};
    `,
  };
};

const ErrorList: React.FC<ErrorListProps> = ({ errors }) => {
  return (
    <List>
      {errors.map((error, index) => (
        <ListItem key={index}>
          <Text style={{ color: "red" }}>{error}</Text>
        </ListItem>
      ))}
    </List>
  );
};

export const useZipFiles = ({ context, theme = "webLight" }: UseZipFilesProps) => {
  const [isDialogOpen, setIsDialogOpen] = React.useState(false);
  const [progress, setProgress] = React.useState(0);
  const [isComplete, setIsComplete] = React.useState(false);
  const [message, setMessage] = React.useState<{ text: string; type: "success" | "error" } | null>(null);
  const [errors, setErrors] = React.useState<string[]>([]);

  const styles = useStyles();

  const isValidSharePointFolderUrl = (folderUrl: string): boolean => {
    const folderUrlPattern = /^\/(sites|teams)\/[\w-]+\/(?:Shared Documents|[\w-]+)(?:\/[\w\-.]+)*\/?$/;
    return folderUrlPattern.test(folderUrl);
  };

  const detectApiType = (url: string): "graph" | "sharepoint" | null => {
    if (url.includes("/drives/") || url.includes("/me/drive/")) {
      return "graph";
    } else if (/^\/(sites|teams)\/[\w-]+\/(?:Shared Documents|[\w-]+)(?:\/[\w\-.]+)*\/?$/.test(url)) {
      return "sharepoint";
    }
    return null;
  };

  const retryFetch = async (url: string, retries: number = MAX_RETRIES): Promise<Response> => {
    let attempt = 0;
    while (attempt < retries) {
      try {
        const response = await fetch(url);
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        return response;
      } catch (error) {
        console.error(`Retry attempt ${attempt + 1} failed for ${url}`);
        attempt++;
        await new Promise((resolve) => setTimeout(resolve, 1000 * Math.pow(2, attempt))); // Exponential backoff
      }
    }
    throw new Error(`Failed to fetch ${url} after ${retries} retries.`);
  };

  const fetchGraphFolderContents = async (folderUrl: string): Promise<IFile[]> => {
    const files: IFile[] = [];
    const fetchGraphFolder = async (folderUrl: string, folderPath: string = ""): Promise<void> => {
      try {
        const msGraphClient = await context.msGraphClientFactory.getClient("3");
        const response = await msGraphClient.api(`${folderUrl}:/children`).version("v1.0").get();

        for (const item of response.value) {
          if (item.folder) {
            const subFolderUrl = `${folderUrl}/${item.name}`;
            const subFolderPath = folderPath ? `${folderPath}/${item.name}` : item.name;
            await fetchGraphFolder(subFolderUrl, subFolderPath);
          } else if (item.file) {
            files.push({
              url: item["@microsoft.graph.downloadUrl"],
              filename: item.name,
              folderName: folderPath,
            });
          }
        }
      } catch (error) {
        console.error(`Error fetching folder contents: ${error.message}`);
        setErrors((prev) => [...prev, `Graph API error: ${error.message}`]);
      }
    };
    await fetchGraphFolder(folderUrl);
    return files;
  };

  const fetchSharePointFolderContents = async (folderUrl: string): Promise<IFile[]> => {
    const files: IFile[] = [];
    try {
      const spHttpClient = context.spHttpClient;
      const response = await spHttpClient.get(
        `${context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')?$expand=Folders,Files`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const data = await response.json();
      const folderFiles = data.Files || [];
      const subfolders = data.Folders || [];

      folderFiles.forEach((file: any) =>
        files.push({
          url: file.ServerRelativeUrl,
          filename: file.Name,
          folderName: folderUrl.replace(/\/$/, ""),
        })
      );
      for (const subfolder of subfolders) {
        const subfolderFiles = await fetchSharePointFolderContents(subfolder.ServerRelativeUrl);
        files.push(...subfolderFiles);
      }
    } catch (error) {
      console.error(`SharePoint API error: ${error.message}`);
      setErrors((prev) => [...prev, `SharePoint API error: ${error.message}`]);
    }
    return files;
  };

  const processDownloads = async (files: IFile[], zip: JSZip) => {
    const queue: Promise<void>[] = [];
    let completed = 0;

    for (const file of files) {
      if (queue.length >= MAX_CONCURRENT_DOWNLOADS) {
        await Promise.race(queue);
      }
      const task = async () => {
        try {
          const response = await retryFetch(file.url);
          const blob = await response.blob();
          const filePath = file.folderName ? `${file.folderName}/${file.filename}` : file.filename;
          zip.file(filePath, blob);
          completed++;
          setProgress(completed / files.length);
        } catch (error) {
          console.error(`Failed to download ${file.filename}: ${error.message}`);
          setErrors((prev) => [...prev, `Failed to download ${file.filename}: ${error.message}`]);
        } finally {
          queue.splice(queue.indexOf(taskPromise), 1); // Safely reference the promise
        }
      };

      const taskPromise = task();
      queue.push(taskPromise);

      await Promise.all(queue); // Ensure all remaining downloads finish
    }
  };

  const createAndDownloadZip = React.useCallback(
    async (filesOrFolderUrl: string | IFile[], fileName: string) => {
      setIsDialogOpen(true);
      setProgress(0);
      setIsComplete(false);
      setMessage(null);
      setErrors([]);

      const zip = new JSZip();
      let files: IFile[] = [];

      if (typeof filesOrFolderUrl === "string") {
        const apiType = detectApiType(filesOrFolderUrl);
        if (apiType === "graph") {
          files = await fetchGraphFolderContents(filesOrFolderUrl);
        } else if (apiType === "sharepoint") {
          if (!isValidSharePointFolderUrl(filesOrFolderUrl)) {
            setMessage({ text: "Invalid SharePoint folder URL format.", type: "error" });
            setIsDialogOpen(false);
            return;
          }
          files = await fetchSharePointFolderContents(filesOrFolderUrl);
        } else {
          setMessage({ text: "Unrecognized folder URL format.", type: "error" });
          setIsDialogOpen(false);
          return;
        }
      } else {
        files = filesOrFolderUrl;
      }
      await processDownloads(files, zip);
      try {
        setMessage({ text: "Creating ZIP file...", type: "success" });
        const content = await zip.generateAsync({ type: "blob" });

        saveAs(content, fileName);
        setIsComplete(true);
        setMessage({ text: "Your ZIP file was created and downloaded!", type: "success" });
      } catch (error) {
        console.error("Error creating ZIP file:", error);
        setErrors((prev) => [...prev, "Error creating ZIP file."]);
      }
    },
    [context]
  );

  const getTheme = (theme: string): Theme => {
    switch (theme) {
      case "webLight":
        return webLightTheme;
      case "webDark":
        return webDarkTheme;
      case "teamsLight":
        return teamsLightTheme;
      case "teamsDark":
        return teamsDarkTheme;
      case "teamsHighContrast":
        return teamsHighContrastTheme;
      default:
        return webLightTheme;
    }
  };

  const ProgressDialog = (
    <IdPrefixProvider value="zip-">
      <FluentProvider theme={getTheme(theme)}>
        <Dialog open={isDialogOpen} onOpenChange={() => setIsDialogOpen(false)} modalType="alert">
          <DialogSurface>
            <DialogBody>
              <DialogTitle>{isComplete ? "Download Complete" : "Creating ZIP File"}</DialogTitle>
              <DialogContent>
                <div  className={styles.container}>
                  {message && <Text  className={message.type === "error" ? styles.errorText : styles.sucessText}>{message.text}</Text>}
                  {progress < 1 && !isComplete && (
                    <>
                      <Field validationMessage={`${Math.round(progress * 100)}%`} validationState="none">
                        <ProgressBar value={progress} />
                      </Field>
                      <Spinner label="Compressing files..." size="medium" labelPosition="above" />
                    </>
                  )}
                  {errors.length && (
                    <div className={styles.errorLogContainer}>
                      <Text>Error Log:</Text>
                      <ErrorList errors={errors} />
                    </div>
                  )}
                </div>
              </DialogContent>
              <DialogActions>
                <DialogTrigger disableButtonEnhancement>
                  <Button appearance="secondary" onClick={() => setIsDialogOpen(false)}>
                    Close
                  </Button>
                </DialogTrigger>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      </FluentProvider>
    </IdPrefixProvider>
  );

  return { createAndDownloadZip, ProgressDialog };
};

export default useZipFiles;
