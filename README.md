# Search O Drive

## Overview
Search O Drive is a GUI application developed for Transportation Compliance Service (TCS) to enhance the efficiency of searching client information within the company's O drive. Designed with simplicity and speed in mind, it significantly reduces the time employees spend looking for client directories compared to traditional search methods.

## Key Features
- **User-Friendly Interface**: Developed with CustomTkinter to ensure a modern and accessible user experience for all TCS employees, regardless of their technical proficiency.
- **Instant Search Results**: Offers practically instantaneous search capabilities, vastly outperforming traditional Windows Explorer searches.
- **Dynamic Search Suggestions**: Features a dynamic list that updates in real-time as the user types, showing companies that match the entered text, facilitating quick and accurate searches.
- **Multi-Directory Search**: Capable of searching across multiple folders simultaneously, displaying whether a client is a Transportation Compliance Service or Employer Solutions client, including directory location.
- **Automatic Updates**: Utilizes a batch script alongside Syspin for seamless updates and program management, including auto-pinning the updated application to the user's taskbar for easy access.

## Technologies Used
- **Python**: Core programming language for developing the application.
- **CustomTkinter**: Used for crafting the graphical user interface.
- **PyInstaller**: Converts the Python script into an executable file for easy distribution and use.
- **Batch Scripting**: Facilitates the automatic update process.
- **Syspin**: A CLI utility that manages pinning the application to the taskbar.

## Challenges and Solutions
- **Learning GUI Development**: The transition to GUI development was initially challenging. Through diligent study and application of CustomTkinter documentation and examples, a modern and intuitive interface was created.
- **Implementing Dynamic Search**: Developing a responsive search suggestion list required innovative use of listbox updates based on user input, significantly enhancing the search efficiency.
- **Directory Loading**: Mapped display names to specific attributes using dictionaries to accurately load the correct company directories upon selection.
- **Deployment and Updates**: Addressed the need for easy deployment and updates among non-technical users by packaging the application with PyInstaller and creating a user-friendly batch script for automated updates.

## Impact and Significance
Search O Drive has become an essential tool for TCS employees, utilized daily for efficient client lookup. It has dramatically reduced search times, streamlined operations, and contributed to a more productive work environment. The application's design and features, such as its auto-minimization functionality, exemplify a commitment to enhancing user experience and operational efficiency at TCS.

## Usage
This application is designed exclusively for TCS employees and is not intended for external use. The code is available for review; however, running the application outside of TCS's environment is not recommended.

## Contributions
As a project developed specifically for Transportation Compliance Service, Search O Drive is not open to external contributions.

## License
The "Search O Drive" project is made available for viewing and download with the understanding that it is specifically designed for use within Transportation Compliance Service (TCS). While the code can be reviewed for educational or informational purposes, running or deploying it outside of the TCS environment is not recommended and may not be practical due to its tailored nature.
