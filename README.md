# Excel File Processor

Excel File Processor is a web application that allows users to upload multiple Excel files, folders, or zip archives, and perform find-and-replace operations across all uploaded files. The processed files are then downloaded as a zip archive.

## Features

- **Multi-File Processing**: Upload multiple Excel files, folders, or zip archives seamlessly.
- **Smart Search & Replace**: Find and replace text across multiple Excel files instantly.
- **Bulk Transformation**: Process multiple files simultaneously with a single click.

## Tech Stack

- **Frontend**: HTML, CSS, JavaScript
- **Backend**: Rust with Actix-web
- **Libraries**: JSZip, SheetJS, calamine, rust_xlsxwriter

![00](https://github.com/user-attachments/assets/1f62069e-8f11-4e62-8fc6-eb3be5a91478)
![01](https://github.com/user-attachments/assets/f37ad217-bf39-451c-937b-96cb6e7903a7)

## Getting Started

### Prerequisites

- Rust and Cargo installed on your machine.
- Node.js and npm installed for frontend dependencies.

### Installation

1. **Clone the repository**:
    ```sh
    git clone https://github.com/your-username/excel-file-processor.git
    cd excel-file-processor
    ```

2. **Backend Setup**:
    ```sh
    cargo build
    cargo run
    ```

3. **Frontend Setup**:
    - Open the `index.html` file in your browser.

### Usage

1. **Start the Server**:
    ```sh
    cargo run
    ```

2. **Open the Application**:
    - Open `index.html` in your browser.
    - Upload Excel files, folders, or zip archives.
    - Enter the text to find and the text to replace.
    - Click on "Process Files" to perform the find-and-replace operation.
    - Download the processed files as a zip archive.

## Project Structure
    excel-file-processor/
    ├── Cargo.toml
    ├── src/
    │   ├── main.rs
    ├── index.html
    ├── README.md


## Contributing

Contributions are welcome! Please open an issue or submit a pull request.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Thanks to the contributors of the libraries used in this project.

## Contact

For any questions or feedback, feel free to reach out to [sunnyshabanali@acm.org](mailto:sunnyshabanali@acm.org).

---

Happy processing! 🚀
