using System;
using Microsoft.Office.Interop.Access.DAO;
using Windows.Media.SpeechRecognition;
/// <summary>
/// Summary description for Class1
/// </summary>
public class Sample
{
    public void Display()
    {
        System.Console.WriteLine("Hello, World!");
    }

    static void main()
    {
    }
    protected async override void OnLaunched(LaunchActivatedEventArgs e)
    {
				try
        {
            StorageFile vcdStorageFile = await Package.Current.InstalledLocation.GetFileAsync(@"AQC.xml");
            await VoiceCommandDefinitionManager.InstallCommandDefinitionsFromStorageFileAsync(vcdStorageFile);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine("There was an error registering the Voice Command Definitions", ex);
        }
    }
    protected override void OnActivated(IActivatedEventArgs e)
    {
        // Handle when app is launched by Cortana
        if (e.Kind == ActivationKind.VoiceCommand)
        {
            VoiceCommandActivatedEventArgs commandArgs = e as VoiceCommandActivatedEventArgs;
            SpeechRecognitionResult speechRecognitionResult = commandArgs.Result;

            string voiceCommandName = speechRecognitionResult.RulePath[0];
            string textSpoken = speechRecognitionResult.Text;
            IReadOnlyList<string> recognizedVoiceCommandPhrases;

            System.Diagnostics.Debug.WriteLine("voiceCommandName: " + voiceCommandName);
            System.Diagnostics.Debug.WriteLine("textSpoken: " + textSpoken);

            switch (voiceCommandName)
            {
                case "Count_Pages":
                    System.Diagnostics.Debug.WriteLine("Count_Pages command");
                    break;

                    // will run C batch files:
                    // system("mybatchfile.bat");  

                    async void Colors_Click(object sender, RoutedEventArgs e)
                    {
                        // Create an instance of SpeechRecognizer.
                        var speechRecognizer = new Windows.Media.SpeechRecognition.SpeechRecognizer();

                        // Add a grammar file constraint to the recognizer.
                        var storageFile = await Windows.Storage.StorageFile.GetFileFromApplicationUriAsync(new Uri("ms-appx:///CountPages.grxml"));
                        var grammarFileConstraint = new Windows.Media.SpeechRecognition.SpeechRecognitionGrammarFileConstraint(storageFile, "rCountPages");

                        speechRecognizer.UIOptions.ExampleText = @"Job Number, ####, pages, ###, single digit numbers 0 - 9";
                        speechRecognizer.Constraints.Add(grammarFileConstraint);

                        // Access the value of {fields} in the voice command.
                        string digit1 = this.SemanticInterpretation("digit1", speechRecognitionResult);
                        string digit2 = this.SemanticInterpretation("digit2", speechRecognitionResult);
                        string digit3 = this.SemanticInterpretation("digit3", speechRecognitionResult);
                        string digit4 = this.SemanticInterpretation("digit4", speechRecognitionResult);
                        string digit5 = this.SemanticInterpretation("digit5", speechRecognitionResult);
                        string digit6 = this.SemanticInterpretation("digit6", speechRecognitionResult);
                        string digit7 = this.SemanticInterpretation("digit7", speechRecognitionResult);

                        string vCourtDatesID = digit1 & digit2 & digit3 & digit4;

                            string vActualQuantity = digit5 & digit6 & digit7;

                            // Compile the constraint.
                        await speechRecognizer.CompileConstraintsAsync();

                        // Start recognition.
                        Windows.Media.SpeechRecognition.SpeechRecognitionResult speechRecognitionResult = await speechRecognizer.RecognizeWithUIAsync();

                        // Do something with the recognition result.
                        var messageDialog = new Windows.UI.Popups.MessageDialog(speechRecognitionResult.Text, "Updated Job Number " & vCourtDatesID & " with Actual Page Count of " & vActualQuantity & ".  Thanks!");
                        await messageDialog.ShowAsync();
                        // DO NOT DO THIS system("\\hubcloud\evoingram\scripts\CountPages.bat");

                        Microsoft.Office.Interop.Access.Application acApp;
                        this.Activate();
                        acApp = (Microsoft.Office.Interop.Access.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Access.Application");
                        Microsoft.Office.Interop.Access.Dao.Database cdb = acApp.CurrentDb();
                        cdb.Execute("UPDATE CourtDates SET ActualQuantity = " & vActualQuantity & " WHERE [CourtDates].[ID] = " & vCourtDatesID & ";");
                    }
            }
        }
    }

    

};