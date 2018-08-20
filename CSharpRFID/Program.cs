using System;
using System.IO.Ports;
using System.Management;

namespace CSharpRFID
{
    class Program
    {
        /// <summary>
        /// O objeto _serialport é utilizado para fazer a conexão e a leitura da porta de comunicação com o Arduino
        /// </summary>
        private static SerialPort _serialPort;

        /// <summary>
        /// Este sistema utiliza a entrada do windows system32 através da leitura em Win32_PnPEntity
        /// Ao fazer a leitura das informações o sistema irá filtrar por informações do fabricante
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            _serialPort = new SerialPort
            {
                // Indica que os bytes nulos serão ignorados durante a transmissão
                DiscardNull = true
            };

            Console.WriteLine("Iniciando o serviço de identificação RFID");

            // Tenta fazer a detecção do hardware
            AutodetectArduinoPort();
            Console.ReadKey();
        }

        /// <summary>
        /// Detecção de informações do Hardware.
        /// Os possíveis tipos de hardware a serem detectados devem ser listados conforme as entradas abaixo (CR340 e Arduino)
        /// Estas informações podem ser obtidas através do painel de controle do windows, dispositivos
        /// </summary>
        static void AutodetectArduinoPort()
        {

            // Você deve fazer referência ao objeto System.Management diretamente através do painel de referências do VS
            ManagementObjectSearcher searcher = VniPnp();

            foreach (ManagementObject item in searcher.Get())
            {
                string Description = item["Description"]?.ToString();
                string Name = item["Name"]?.ToString();
                string SystemName = item["SystemName"]?.ToString();

                if (Description != null && Name.Contains("(COM"))
                {
                    if (Description.Contains("CH340") || Description.Contains("Arduino"))
                    {
                        // Hardware encontrado. Armazenando as inforações
                        Console.WriteLine($"Encontrado o leitor RFID em {Name}");
                        HardwareInfo LeitorRFIDInfo = new HardwareInfo
                        {
                            PortName = Name.Substring(Name.IndexOf("(COM")).Replace("(", "").Replace(")", ""),
                            Description = Description,
                            SystemName = SystemName
                        };
                        Console.WriteLine(LeitorRFIDInfo.ToString());

                        _serialPort.PortName = LeitorRFIDInfo.PortName;
                        _serialPort.DataReceived += SerialPort_DataReceived;
                        _serialPort.ErrorReceived += SerialPort_ErrorReceived;

                        // Inicia processo de conexão serial
                        _serialPort.Open();
                        Console.WriteLine($"Situação: {_serialPort.IsOpen}");
                        return;
                    }
                }
            }
            
            string mensagem = "Não foi possível identificar o equipamento RFID.\r\nVerifique se os drives estão instalados e se o equipamento encontra-se corretamente conectado ao computador";
            Console.WriteLine(mensagem);

        }

        /// <summary>
        /// Abre uma conexão de leitora PNP do windows
        /// </summary>
        /// <returns></returns>
        static ManagementObjectSearcher VniPnp()
        {
            ManagementScope connectionScope = new ManagementScope();
            SelectQuery serialQuery = new SelectQuery("SELECT * FROM Win32_PnPEntity");
            return new ManagementObjectSearcher(connectionScope, serialQuery);
        }

        /// <summary>
        /// Ao encontrar algum erro de conexão ou leitura
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        static void SerialPort_ErrorReceived(object sender, SerialErrorReceivedEventArgs e)
        {
            Console.WriteLine($"Erro do sistema serial {e.ToString()}");
        }

        /// <summary>
        /// Recebe as informações do leitor RFID
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        static void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            Console.WriteLine($"Serial recebido");
            string RxString = _serialPort.ReadExisting();
            string SerialID;

            if (RxString.IndexOf("\r\n") > -1 || RxString.IndexOf("|") > -1)
            {
                SerialID = RxString.Replace("\r\n", "").Replace("|", "");
                Console.WriteLine($"Serial: {SerialID}");
            }
        }
    }

    /// <summary>
    /// Utilizado para armazenar informações do hardware
    /// </summary>
    internal class HardwareInfo
    {
        public string PortName { get; internal set; }
        public string Description { get; internal set; }
        public string SystemName { get; internal set; }

        public override string ToString() {
            return $" SystemName: {SystemName}\r\n Description:{Description}\r\n PortName:{PortName}";
        }  
        
    }
}
