# Introducing `useZipFiles`: A Powerful React Hook for Creating and Downloading ZIP Files

Managing file downloads and creating ZIP archives can be challenging, especially when dealing with APIs like SharePoint or Microsoft Graph. To simplify this process, I’ve built a versatile React hook, `useZipFiles`, which allows you to fetch files or folders, compress them into a ZIP file, and trigger a download—all in one streamlined package.

---

## 🚀 Features of `useZipFiles`

1. **Supports SharePoint and Microsoft Graph APIs**: Fetch files and folders seamlessly from both APIs.
2. **Recursive Folder Handling**: Automatically fetches nested folders and their contents.
3. **Download Progress Dialog**: Displays download progress with Fluent UI components.
4. **Retry Mechanism**: Handles download failures with a retry strategy.
5. **Error Logging**: Logs errors in an easy-to-read format.

---

## 📂 Where to Find This Hook

The full source code for the `useZipFiles` hook is available on GitHub. You can clone or download the repository here:

**GitHub Repository**: [https://github.com/joaojmendes/useZipFiles](https://github.com/joaojmendes/useZipFiles)

You can copy the source code and use it in your projects. this project has SPFx 1.20 webpart that shows how to use the hook.

---

## 🛠️ How to Install

1. Clone the repository from GitHub:

   ```bash
   git clone https://github.com/joaojmendes/useZipFiles.git
