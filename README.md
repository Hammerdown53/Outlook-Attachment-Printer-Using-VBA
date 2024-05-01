# Outlook-Attachment-Printer-Using-VBA
This VBA script enables users to print selected attachments from Outlook emails. It includes functionality to prompt the user for confirmation before printing and clears temporary folders where attachments are temporarily saved.

## Table of Contents

- [Introduction](#introduction)
- [How to Use](#how-to-use)
- [Contributing](#contributing)
- [License](#license)
- [Credits](#credits)

## Introduction

The script consists of two main procedures:

1. **PrintSelectedAttachments:** This subroutine is triggered when the user wants to print selected attachments from Outlook emails. It prompts the user for confirmation before proceeding to print the attachments.

2. **PrintAttachments:** This subroutine is called by PrintSelectedAttachments to print the selected attachments. It iterates through each attachment, saves it to a temporary folder, and then prints it using the default application associated with the file type.

Additionally, there's a **ClearTempFolders** subroutine provided to clean up temporary folders used to store attachments.

## How to Use

1. Open Outlook and navigate to the email containing the attachments you want to print.
2. Select the email(s) and run the **PrintSelectedAttachments** subroutine. You will be prompted for confirmation.
3. If you confirm, the script will save the attachments to a temporary folder and print them using the default application associated with their file types.

## Contributing

Contributions to this project are welcome. If you have suggestions for improvements or would like to report a bug, please open an issue or submit a pull request.

## License

This project is licensed under the [MIT License](LICENSE).

## Credits

This script was created by Trey McBride.

Feel free to customize and improve upon this script as needed for your use case. Enjoy printing your Outlook attachments!
