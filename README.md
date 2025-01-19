Easy Video Graphics Creation Project Requirements
Austin Black
Jan. 19, 2025
Version 1.0

---

Introduction

The purpose of this document is to outline the features and operating environment for the Automated Video Graphics Creation Project (EVGC). This document defines the features of a minimum viable product (MVP) while leaving room for future development and enhancements.

---

General Description

The EVGC tool is designed to streamline the creation of in-video graphics, particularly lower-thirds, commonly used in live-streaming and video production. The current implementation consists of two main components:

1. Command-Line Interface (CLI):
    - Uses the `python-pptx` library to process PowerPoint (.pptx) files.
    - Extracts text content from slides and generates a new .pptx file formatted as lower-thirds graphics.
    - Outputs a folder containing:
        - The formatted .pptx file with slides customized using assets from an `Assets/` directory.
        - A text (.txt) metadata file corresponding to each slide.
    - Each slide includes three core elements:
        - A background image.
        - A logo image.
        - A text box.

    - Current Limitations:
        - The CLI is not user-friendly.
        - Users cannot easily customize the format of lower-thirds graphics.
        - No preview functionality is available for users to visualize their configurations.

2. PowerPoint VBA Script Macros:
    - Utilizes metadata text files to apply animations to PowerPoint objects created in the CLI step.
    - This component will not be a primary focus for this project.

Primary Objective:
To develop a functional graphical user interface (GUI) for the CLI portion of the EVGC tool. The GUI will improve usability by providing an intuitive interface for input, customization, and previewing.

---

Planned GUI Features

1. File Management:
    - Select input files (“must” support .pptx files).
    - Choose output directories for generated files.

2. Customization Options:
    - Change the images used for lower-thirds graphics (backgrounds, logos, etc.).
    - Adjust the position of text and graphics, maintaining consistency across all slides.
    - Modify background colors for upstream, downstream, and chroma keying compatibility.

3. Preview Functionality:
    - Display a preview image or sample of the current configuration.
    - Update previews in real-time or with a “refresh” button.

---

Functional Requirements

- User Experience:
    - Allow users to quickly create on-screen graphics tailored to specific events.
    - Provide a straightforward and intuitive interface for all features.

- Performance:
    - The GUI must respond immediately to user interactions, with minimal delay.
    - Generating graphics should complete in under five minutes.

- Error Handling:
    - Handle invalid inputs gracefully.
    - Offer informative error messages to assist users in troubleshooting.

---

Non-Functional Requirements

- The tool should maintain high reliability and stability during operation.
- The codebase must be modular and maintainable, allowing for future expansions.
- Adhere to best practices for accessibility and usability in GUI design.

---

Future Goals

1. Combine the two portions into a single program.
2. Support additional input formats beyond .pptx.
3. Expand customization capabilities to include more granular control over graphic elements.

---

This document serves as a foundational guide for the EVGC project, setting clear objectives and expectations for its development.
