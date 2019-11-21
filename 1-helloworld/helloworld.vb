Imports SystemModule Module1
    'This program will display Hello World 
    Sub Main()
        Console.WriteLine("Hello World")
        Console.WriteLine(vbCrLf + "What is your name? ")
        Dim name = Console.ReadLine()
        Dim currentDate = DateTime.Now
        Console.WriteLine($"{vbCrLf}Hello, {name}, on {currentDate:d} at {currentDate:t}!")
        Console.Write(vbCrLf + "Press any key to exit... ")
        Console.ReadKey(True)
    End SubEnd Module
