using System;
using System.Linq;
using System.Text;
using LocalDocumentCompare;

internal static class LanguageDetectionTests
{
    private const string English = "The present invention relates to a method for controlling a motor. The apparatus comprises a housing, a controller and a sensor. The controller receives the signal from the sensor and adjusts the speed of the motor. This arrangement improves the operation of the device. ";
    private const string German = "Die vorliegende Erfindung betrifft ein Verfahren zur Steuerung eines Motors. Die Vorrichtung umfasst ein Gehäuse, eine Steuereinheit und einen Sensor. Die Steuereinheit empfängt das Signal des Sensors und stellt die Drehzahl des Motors ein. Diese Anordnung verbessert den Betrieb der Vorrichtung. ";
    private const string French = "La présente invention concerne un procédé de commande d'un moteur. Le dispositif comprend un boîtier, un contrôleur et un capteur. Le contrôleur reçoit le signal du capteur et ajuste la vitesse du moteur. Cet agencement améliore le fonctionnement du dispositif. ";
    private const string Spanish = "La presente invención se refiere a un método para controlar un motor. El aparato comprende una carcasa, un controlador y un sensor. El controlador recibe la señal del sensor y ajusta la velocidad del motor. Esta disposición mejora el funcionamiento del dispositivo. ";
    private const string Italian = "La presente invenzione riguarda un metodo per controllare un motore. Il dispositivo comprende un alloggiamento, un controllore e un sensore. Il controllore riceve il segnale dal sensore e regola la velocità del motore. Questa disposizione migliora il funzionamento del dispositivo. ";
    private const string Dutch = "De onderhavige uitvinding heeft betrekking op een werkwijze voor het besturen van een motor. De inrichting omvat een behuizing, een besturingseenheid en een sensor. De besturingseenheid ontvangt het signaal van de sensor en past het toerental van de motor aan. ";
    private const string Russian = "Настоящее изобретение относится к способу управления двигателем. Устройство содержит корпус, контроллер и датчик. Контроллер получает сигнал от датчика и регулирует скорость двигателя. Такое расположение улучшает работу устройства. ";
    private const string Japanese = "本発明は、モーターを制御する方法に関する。装置は、ハウジング、制御装置及びセンサーを備える。制御装置はセンサーからの信号を受信し、モーターの速度を調整する。この構成により、装置の動作が改善される。";
    private const string Chinese = "本发明涉及一种控制电机的方法。该装置包括外壳、控制器和传感器。控制器接收来自传感器的信号并调整电机的速度。这种配置改善了设备的运行。";
    private static int checks;

    private static string Repeat(string value, int count) { return String.Concat(Enumerable.Repeat(value, count)); }

    private static void Expect(string label, string text, string status, string code)
    {
        SourceLanguage result = LanguageDetection.Detect(text);
        if (result.Status.ToString() != status || result.Code != code)
            throw new Exception(label + ": expected " + status + "/" + code + "; got " + result.Status + "/" + result.Code);
        checks++;
    }

    private static void ExpectUncertain(string label, string text)
    {
        SourceLanguage result = LanguageDetection.Detect(text);
        if (result.Status == LanguageStatus.Known)
            throw new Exception(label + ": expected Unknown/Mixed; got " + result.Status + "/" + result.Code);
        checks++;
    }

    public static int Main()
    {
        try
        {
            Expect("English", Repeat(English, 10), "Known", "en");
            Expect("German", Repeat(German, 10), "Known", "de");
            Expect("French", Repeat(French, 10), "Known", "fr");
            Expect("Spanish", Repeat(Spanish, 10), "Known", "es");
            Expect("Italian", Repeat(Italian, 10), "Known", "it");
            Expect("Dutch", Repeat(Dutch, 10), "Known", "nl");
            Expect("Russian", Repeat(Russian, 10), "Known", "ru");
            Expect("Japanese", Repeat(Japanese, 14), "Known", "ja");
            Expect("Chinese", Repeat(Chinese, 20), "Known", "zh");
            Expect("NFC normalization", Repeat(French, 10).Normalize(NormalizationForm.FormD), "Known", "fr");
            Expect("Empty", "", "Unknown", "");
            Expect("Null", null, "Unknown", "");
            Expect("Short", English, "Unknown", "");
            Expect("Digits", Repeat("12345 67890 ", 200), "Unknown", "");
            Expect("Repeated labels", Repeat("Claims Description Summary Figure Motor Controller ", 60), "Unknown", "");
            Expect("Malformed Unicode", Repeat(English + "\ud800", 10), "Unknown", "");
            Expect("English/German", Repeat(English, 14) + Repeat(German, 14), "Mixed", "");
            Expect("English/French", Repeat(English, 14) + Repeat(French, 14), "Mixed", "");
            Expect("German/French", Repeat(German, 14) + Repeat(French, 14), "Mixed", "");
            Expect("English/German/French", Repeat(English, 12) + Repeat(German, 12) + Repeat(French, 12), "Mixed", "");
            Expect("English/Japanese", Repeat(English, 14) + Repeat(Japanese, 40), "Mixed", "");
            Expect("Distributed samples", Repeat(English, 200) + Repeat(German, 200), "Mixed", "");
            ExpectUncertain("Translated opening", Repeat(German, 5) + Repeat(English, 200));
            ExpectUncertain("Translated ending", Repeat(English, 200) + Repeat(German, 5));
            ExpectUncertain("Interleaved English/German", Repeat(English + German, 20));
            ExpectUncertain("Interleaved English/French", Repeat(English + French, 20));
            ExpectUncertain("Interleaved German/French", Repeat(German + French, 20));
            if (LanguageDetection.NormalizeCode("en-GB") != "en" || LanguageDetection.NormalizeCode("en-US") != "en" ||
                LanguageDetection.NormalizeCode("zh-Hant") != "zh" || LanguageDetection.NormalizeCode("zh-Hans") != "zh")
                throw new Exception("Regional/script language normalization failed.");
            checks++;
            Console.WriteLine("PASS: " + checks + " local language checks using native Windows ELS; no document fixtures, text logs or network services.");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error.Message); return 1; }
    }
}
