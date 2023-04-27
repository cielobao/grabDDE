# grabDDE
Lets you communicate between different applications running on Windows using DDE (Dynamic Data Exchange).

The script defines a DDE class that encapsulates various DDE-related functions from the user32 library. It also defines a DDEClient class that represents a DDE client and allows the user to perform various operations on it such as establishing a conversation with a DDE server, sending requests and commands, and starting/stopping data advises.

When the server sends data advise, the script uses a _callback function that performs actions when the data advise is received. You can customize this function as needed.

The script also has a WinMSGLoop function that runs continuously and processes messages from the Windows message queue. This function ensures that the data advise callback function is called when data is received from the server.

It's important to note that this script is only for Windows and requires certain libraries to be installed.
