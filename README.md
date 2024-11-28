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

# Rust Project Setup Guide

## Step 1: Download and Install Rust Rover
1. **Download Rust Rover**:
   * Go to the **JetBrains Rust Rover download page**
   * Download and install the appropriate version for your operating system

## Step 2: Download and Install LLVM
1. **Download LLVM**:
   * Go to the **LLVM 1.8.1 release page**
   * Download the appropriate version for your operating system
   * Follow the installation instructions provided on the release page

## Step 3: Set Up a Rust Project in Rust Rover
1. **Open Rust Rover**:
   * Launch Rust Rover

2. **Create a New Rust Project**:
   * Go to `File` > `New` > `Project`
   * Select `Rust` and follow the prompts to create a new Rust project
   * When prompted, download and install `rustup` and its linker as recommended

## Step 4: Copy Backend Files
1. **Copy Cargo.toml**:
   * Open the `Cargo.toml` file in your new Rust project
   * Replace the contents starting from `[dependencies]` to the end with the provided `Cargo.toml` content

2. **Copy main.rs**:
   * Open the `src/main.rs` file in your new Rust project
   * Replace the contents with the provided `main.rs` content

## Step 5: Run the Backend Server
1. **Build and Run the Project**:
   * In Rust Rover, click on the `Run` button or use the terminal to run `cargo run`
   * This will download the required dependencies, build the project, and start the backend server locally

## Step 6: Open the Frontend
1. **Open index.html**:
   * Navigate to the `frontend` directory
   * Open `index.html` in your web browser

## Step 7: Process Excel Files
1. **Use the Frontend**:
   * Follow the instructions on the frontend to upload Excel files
   * Enter find and replace strings
   * Process the files
   * The backend server will handle the file processing and return the processed files as a downloadable ZIP archive

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
