using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Speech.Synthesis;

namespace Apress.VisualCSharpRecipes.Chapter08
{
    class Recipe08_19
    {
        static void Main(string[] args)
        {
            // create a new synthesizer
            SpeechSynthesizer mySynth = new SpeechSynthesizer();

            Console.WriteLine("--- Start of voices list ---");
            foreach (InstalledVoice voice in mySynth.GetInstalledVoices())
            {
                Console.WriteLine("Voice: {0}", voice.VoiceInfo.Name);
                Console.WriteLine("Gender: {0}", voice.VoiceInfo.Gender);
                Console.WriteLine("Age: {0}", voice.VoiceInfo.Age);
                Console.WriteLine("Culture: {0}", voice.VoiceInfo.Culture);
                Console.WriteLine("Description: {0}", voice.VoiceInfo.Description);
            }
            Console.WriteLine("--- End of voices list ---");

            while (true)
            {
                Console.WriteLine("Enter string to speak");
                mySynth.Speak(Console.ReadLine());
                Console.WriteLine("Completed");
            }
        }
    }
}
