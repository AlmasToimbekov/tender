# Tender Excel Add-in

This repository contains the source code for the "Tender" Excel add-in. The add-in allows users to fetch and populate company information in an Excel spreadsheet based on selected IDs.

## Features

- Fetch company information from an external API.
- Populate the fetched data into the Excel spreadsheet.
- Auto-fit column widths for certain columns.
- Set fixed width and enable text wrapping for specific columns.

## Getting Started

### Prerequisites

- [Node.js](https://nodejs.org/) (for development and testing)
- [Office Add-ins](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins) (for Excel)

### Installation

1. **Clone the repository:**

   ```bash
   git clone https://github.com/AlmasToimbekov/tender.git
   cd tender
   ```

1. **Install dependencies:**

   ```bash
   npm install
   ```

## Usage

1. **Select the rows in Excel where the IDs are located.**
1. **Click the "Получить данные" button in the task pane.**
1. **The add-in will fetch the data and populate the columns to the right of the selected range.**
1. **Sideload the add-in in Excel for web:**

   - Open an Excel file on your [web](https://www.microsoft365.com/launch/Excel).
   - [Insert Office Add-ins](https://support.microsoft.com/en-us/office/insert-office-add-ins-into-excel-for-the-web-3a19321c-182e-4cb7-9379-7f646fa2152c) into Excel for the web using the [manifest.xml](./manifest.xml) file

1. **Sideload the Add-in in Excel for Desktop:**

   To be investigated

## Development

### Running Locally & Debugging

1. **Use localhost:**

   - Fix the `manifest.xml` file to use `localhost` addresses.

1. **Start the local web server:**

     Press F5 according to [VS Code debugging](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/debug-office-add-ins-in-visual-studio) instructions

## Deploy
### Hosting on GitHub Pages

1. **Enable GitHub Pages:**

   - Go to your repository settings on GitHub.
   - Scroll down to the **GitHub Pages** section.
   - Under **Source**, select the `gh-pages` branch and the root folder.
   - Click **Save**.

1. **Update Manifest File:**

   Ensure that the URLs in your `manifest.xml` file point to the correct locations on GitHub Pages. Change the version if necessary.

1. **Build and Deploy:**

   ```bash
   npm run build
   npm run deploy
   ```
   That creates/updates the `gh-pages` branch on GitHub, that serves your add-in.

1. Follow the `Usage` section.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.