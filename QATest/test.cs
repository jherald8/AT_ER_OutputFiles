
using System;
using System.Text;

Console.WriteLine("My Program"); // Enter starting breakpoint here

string userChoice;
    do
    {
        // Type your username and press enter
        Console.WriteLine("Enter username:");

        // Create a string variable and get user input from the keyboard and store it in the variable
        string userName = Console.ReadLine();

        // Type your pasword and press enter
        Console.WriteLine("Enter password:");

        string password = MaskInput();

        Console.WriteLine();

        // Re-enter pasword and press enter
        Console.WriteLine("Re-enter pasword:");
        string password2 = MaskInput();


        // Compare if password and the re-entered password is matched
        if (password == password2)
        {
            Console.WriteLine("Congratulations " + userName + "! password matched");
        }
        else
        {
            Console.WriteLine("I'm sorry " + userName + "! password is not matched");
        }


        // Function to read input and mask it
        static string MaskInput()
        {
            string input = string.Empty;

            while (true)
            {
                var keyInfo = Console.ReadKey(intercept: true); // Read key without displaying it

                // Check for Enter key
                if (keyInfo.Key == ConsoleKey.Enter)
                {
                    // Exit the loop if Enter is pressed
                    break;
                }
                else if (keyInfo.Key == ConsoleKey.Backspace)
                {
                    // Handle backspace
                    if (input.Length > 0)
                    {
                        input = input.Substring(0, input.Length - 1);
                        Console.Write("\b \b"); // Remove the last character on the console
                    }
                }
                else
                {
                    input += keyInfo.KeyChar; // Append
                                              // the character to the input string
                    Console.Write('*'); // Display an asterisk instead of the character
                }
            }

            Console.WriteLine();
            return input; // Return the original input

        }
        Console.Write("Do you want to perform the task again? (yes/no): ");
        userChoice = Console.ReadLine().ToLower(); // Read input and convert to lowercase

    }

    while (userChoice == "yes"); // Repeat if user types 'yes'

    Console.WriteLine("Exiting the program. Goodbye!"); // Enter ending  breakpoint here



