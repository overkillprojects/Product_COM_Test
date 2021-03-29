using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.IO.Ports;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace Product_COM_Test
{
    static class Constants
    {
        public const int USB_EXTERNAL_BUFFER_SIZE = 64;
        public const int USB_INTERNAL_BUFFER_SIZE = 4096;
        public const int SERIAL_NUMBER_SIZE = 12;
        public const int Files_VERSION_SIZE = 4;
        public const int RESET_Files_REPLY_SIZE = 2;
        public const int FEED_REPLY_SIZE = 2;
    }
    public partial class MainWindow : Window
    {
        static SerialPort _serialPort = new SerialPort();
        static OpenFileDialog openDialog = new OpenFileDialog();
        static FileStream sample_file;
        static Encoding enc8 = Encoding.UTF8;
        static byte activity_file = 0;
        static byte activity_current = 0;
        static byte activity_previous = 0;
        static int send_data_size = 0;
        static int data_buffers_sent = 0;
        static UInt32 Files_version = 0;

        static byte[] read_buffer = new byte[Constants.USB_INTERNAL_BUFFER_SIZE];
        static byte[] send_buffer = new byte[Constants.USB_INTERNAL_BUFFER_SIZE * 7];
        static byte[] getSerial = { 0x41, 0x50, 0x52, 0x4F, 0x44, 0x55, 0x43, 0x54,
            0x00, 0x00,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF };
        static byte[] getFSVersion = { 0x41, 0x50, 0x52, 0x4F, 0x44, 0x55, 0x43, 0x54, 0x00, 0x01,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF };
        static byte[] setFSVersion = { 0x41, 0x50, 0x52, 0x4F, 0x44, 0x55, 0x43, 0x54, 0x00, 0x02,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF };
        static byte[] resetFiles = { 0x41, 0x50, 0x52, 0x4F, 0x44, 0x55, 0x43, 0x54, 0xDE, 0xAD,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF };
        static byte[] sendSample = { 0x41, 0x50, 0x52, 0x4F, 0x44, 0x55, 0x43, 0x54, 0xFE, 0xED,
            0x00, 0x01, 0x10, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF };

        static bool comPortsListed;

        private enum ProductCOMEvents
        {
            None,
            GetSerialNumberDone,
            getFSVersionDone,
            setFSVersionDone,
            resetFilesDone,
            FeedDone,
        }

        static ProductCOMEvents event_type = ProductCOMEvents.None;

        public MainWindow()
        {
            InitializeComponent();
            disableFields();
            SetComPortComboBoxDefaults();
        }

        private void comboBoxPorts_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Since we loaded the combobox with ComPortProperties earlier, we need to cast from object to ComPortProperties
            // to get the correct value out. The SelectedItem can be null if there are no entries in the combobox
            ComPortProperties properties = (ComPortProperties) comboBoxPorts.SelectedItem;
            if(properties == null)
            {
                return;
            }

            _serialPort.PortName = properties.ComPortName;
            _serialPort.BaudRate = 115200;
            _serialPort.Parity = Parity.None;
            _serialPort.DataBits = 8;
            _serialPort.StopBits = StopBits.One;
            _serialPort.Handshake = Handshake.None;
            _serialPort.ReadBufferSize = 4096;
            _serialPort.WriteTimeout = 500;
            _serialPort.ReadTimeout = 5000;
            _serialPort.DtrEnable = true; //enable Data Terminal Ready
            _serialPort.RtsEnable = true; //enable Request to send
        }

        private void buttonOpenPort_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_serialPort.IsOpen) _serialPort.Close();
                if (_serialPort != null)
                {
                    _serialPort.Open();
                    _serialPort.DataReceived += Serial_DataReceived;
                    _serialPort.Read(read_buffer, 0, 2 * Constants.USB_EXTERNAL_BUFFER_SIZE);
                    enableFields();
                    this.Dispatcher.Invoke((Action)(() =>
                    {
                        label1.Content = "Connected to " + _serialPort.PortName;
                    }));
                }
                else
                {
                    this.Dispatcher.Invoke((Action)(() =>
                    {
                        label1.Content = "Choose a COM port.";
                    }));
                }
            }
            catch (Exception ex)
            {
                this.Dispatcher.Invoke((Action)(() =>
                {
                    label1.Content = "Invalid COM port.";
                }));
            }
        }

        private void Serial_DataReceived(object s, SerialDataReceivedEventArgs e)
        {
            switch (event_type)
            {
                case ProductCOMEvents.GetSerialNumberDone:
                    {
                        try
                        {
                            readData(Constants.USB_EXTERNAL_BUFFER_SIZE, Constants.SERIAL_NUMBER_SIZE, labelSerialNumber);
                        }
                        catch (Exception ex)
                        {
                            onError(ex);
                        }
                        event_type = ProductCOMEvents.None;
                        break;
                    }
                case ProductCOMEvents.getFSVersionDone:
                    {
                        try
                        {
                            _serialPort.Read(read_buffer, 0, Constants.Files_VERSION_SIZE + 1);
                            string ret = "";
                            Files_version = 0;
                            for (int i = 0; i < Constants.Files_VERSION_SIZE; i++)
                            {
                                ret += read_buffer[i].ToString("X") + " ";
                                Files_version += read_buffer[i];
                            }
                            this.Dispatcher.Invoke((Action)(() =>
                            {
                                labelFilesVersion.Content = ret;
                            }));
                            if (Files_version == 0xFFFFFFFF)
                                Files_version = 0;
                            else
                                ++Files_version;
                        }
                        catch (Exception ex)
                        {
                            onError(ex);
                        }
                        event_type = ProductCOMEvents.None;
                        break;
                    }
                case ProductCOMEvents.setFSVersionDone:
                    {
                        try
                        {
                            readData(Constants.USB_EXTERNAL_BUFFER_SIZE, Constants.RESET_Files_REPLY_SIZE, labelSetVersion);
                        }
                        catch (Exception ex)
                        {
                            onError(ex);
                        }
                        event_type = ProductCOMEvents.None;
                        break;
                    }
                case ProductCOMEvents.resetFilesDone:
                    {
                        try
                        {
                            readData(Constants.USB_EXTERNAL_BUFFER_SIZE, Constants.RESET_Files_REPLY_SIZE, labelResetFiles);
                            setFSVersion[10] = (byte)((Files_version >> 24) & 0xFF);
                            setFSVersion[11] = (byte)((Files_version >> 16) & 0xFF);
                            setFSVersion[12] = (byte)((Files_version >> 8) & 0xFF);
                            setFSVersion[13] = (byte)((Files_version >> 0) & 0xFF);
                            event_type = ProductCOMEvents.setFSVersionDone;
                            _serialPort.Write(setFSVersion, 0, Constants.USB_EXTERNAL_BUFFER_SIZE);
                        }
                        catch (Exception ex)
                        {
                            onError(ex);
                        }
                        break;
                    }
                case ProductCOMEvents.FeedDone:
                    {
                        try
                        {
                            if (data_buffers_sent == 0)
                            {
                                readData(Constants.USB_EXTERNAL_BUFFER_SIZE, Constants.FEED_REPLY_SIZE, labelSendSample);

                                /* Send dummy frame to clear double buffer */
                                byte[] dummy_buffer = new byte[64];
                                _serialPort.Write(dummy_buffer, 0, Constants.USB_EXTERNAL_BUFFER_SIZE);



                                sample_file.Read(send_buffer, 0, send_data_size);

                                /* Now send to CDC */
                                _serialPort.Write(send_buffer, data_buffers_sent++, Constants.USB_EXTERNAL_BUFFER_SIZE);
                                sample_file.Close();
                            }
                            else if (send_data_size >= data_buffers_sent * Constants.USB_EXTERNAL_BUFFER_SIZE)
                            {
                                readData(Constants.USB_EXTERNAL_BUFFER_SIZE, Constants.FEED_REPLY_SIZE, labelB469);
                                _serialPort.Write(send_buffer, data_buffers_sent++, Constants.USB_EXTERNAL_BUFFER_SIZE);
                            }
                            else
                            {
                                event_type = ProductCOMEvents.None;
                                sample_file.Close();
                                this.Dispatcher.Invoke((Action)(() =>
                                {
                                    textBoxSampleFile.Text = "";
                                }));
                            }
                        }
                        catch (Exception ex)
                        {
                            onError(ex);
                        }

                        break;
                    }
                default:
                    {
                        try
                        {
                            readData(Constants.USB_EXTERNAL_BUFFER_SIZE, Constants.USB_EXTERNAL_BUFFER_SIZE, label1);
                        }
                        catch (Exception ex)
                        {
                            onError(ex);
                        }
                        event_type = ProductCOMEvents.None;
                        break;
                    }
            }
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            if (_serialPort.IsOpen)
            {
                _serialPort.Close();
            }
            if (_serialPort != null)
            {
                _serialPort.Dispose();
            }
        }

        private void buttonClosePort_Click(object sender, RoutedEventArgs e)
        {
            if (_serialPort.IsOpen)
            {
                _serialPort.Close();
                _serialPort.Dispose();
            }
            _serialPort = new SerialPort();
            comboBoxPorts.SelectedIndex = -1;
            disableFields();
            this.Dispatcher.Invoke((Action)(() =>
            {
                label1.Content = "Choose a COM port.";
            }));
        }

        private void buttonGetSerial_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_serialPort.IsOpen)
                {
                    event_type = ProductCOMEvents.GetSerialNumberDone;
                    _serialPort.Write(getSerial, 0, Constants.USB_EXTERNAL_BUFFER_SIZE);
                }
            }
            catch (Exception ex)
            {
                _serialPort.Close();
                this.Dispatcher.Invoke((Action)(() =>
                {
                    labelSerialNumber.Content = ex.Message;
                }));
            }
        }

        private void buttonGetFilesVersion_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_serialPort.IsOpen)
                {
                    this.Dispatcher.Invoke((Action)(() =>
                    {
                        event_type = ProductCOMEvents.getFSVersionDone;
                    }));
                    _serialPort.Write(getFSVersion, 0, Constants.USB_EXTERNAL_BUFFER_SIZE);
                }
            }
            catch (Exception ex)
            {
                _serialPort.Close();
                this.Dispatcher.Invoke((Action)(() =>
                {
                    labelFilesVersion.Content = ex.Message;
                }));
            }
        }

        private void buttonResetFiles_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_serialPort.IsOpen)
                {
                    event_type = ProductCOMEvents.resetFilesDone;
                    _serialPort.Write(resetFiles, 0, Constants.USB_EXTERNAL_BUFFER_SIZE);
                }
            }
            catch (Exception ex)
            {
                _serialPort.Close();
                this.Dispatcher.Invoke((Action)(() =>
                {
                    labelResetFiles.Content = ex.Message;
                }));
            }
        }

        private void readData(int buff_size, int reply_size, Label disp_label)
        {
            try
            {
                _serialPort.Read(read_buffer, 0, reply_size + 1);
                string ret = "";
                for (int i = 0; i < reply_size; i++)
                {
                    ret += read_buffer[i].ToString("X") + " ";
                }
                this.Dispatcher.Invoke((Action)(() =>
                {
                    disp_label.Content = ret;
                }));
            }
            catch (Exception ex)
            {
                onError(ex);
            }
        }

        private void onError(Exception ex)
        {
            _serialPort.Close();
            disableFields();
            this.Dispatcher.Invoke((Action)(() =>
            {
                label1.Content = "ERROR: " + ex.Message;
            }));
            event_type = ProductCOMEvents.None;
            if (sample_file != null)
            {
                sample_file.Close();
            }
        }

        private void clearBuffer(Array buffer)
        {
            Array.Clear(buffer, 0, Constants.USB_INTERNAL_BUFFER_SIZE);
        }

        private void disableFields()
        {
            this.Dispatcher.Invoke((Action)(() =>
            {
                buttonClosePort.IsEnabled = false;
                buttonOpenPort.IsEnabled = true;
                comboBoxPorts.IsEnabled = true;
                buttonGetSerial.IsEnabled = false;
                buttonGetFSVersion.IsEnabled = false;
                buttonResetFiles.IsEnabled = false;

                labelFilesVersion.Content = "";
                labelSerialNumber.Content = "";
                labelResetFiles.Content = "";
            }));
        }

        private void enableFields()
        {
            this.Dispatcher.Invoke((Action)(() =>
            {
                comboBoxPorts.IsEnabled = false;
                buttonClosePort.IsEnabled = true;
                buttonOpenPort.IsEnabled = false;

                buttonGetSerial.IsEnabled = true;
                buttonGetFSVersion.IsEnabled = true;
                buttonResetFiles.IsEnabled = true;

                labelFilesVersion.Content = "";
                labelSerialNumber.Content = "";
                labelResetFiles.Content = "";
            }));
        }

        private void buttonOpenFile_Click(object sender, RoutedEventArgs e)
        {
            openDialog.Multiselect = false;
            openDialog.Filter = "WAV (PCM) files (*.wav)|*.wav|All files (*.*)|*.*";
            openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyMusic);
            if (openDialog.ShowDialog() == true)
            {
                textBoxSampleFile.Text = openDialog.FileName;
            }
        }

        private void buttonSendSample_Click(object sender, RoutedEventArgs e)
        {
            string data_test;
            int i = 0;

            /* Get Activity and File */
            activity_current = byte.Parse(textBoxActivity.Text);
            if (activity_current != activity_previous) activity_file = 0;

            /* Open file if it isn't open */
            try
            {
                sample_file = File.OpenRead(openDialog.FileName);
            }
            catch (Exception ex) { }

            /* Search for "data" */
            sample_file.Read(read_buffer, 0, 3);
            do
            {
                sample_file.Read(read_buffer, 3 + i, 1);
                data_test = enc8.GetString(read_buffer, i++, 4);
            } while (data_test != "data");

            /* Now read next 4 bytes for file size */
            sample_file.Read(read_buffer, 0, 4);
            send_data_size = BitConverter.ToInt32(read_buffer, 0);

            /* Now read that many bytes into the buffer to send */
            int num_bytes = 4096;
            this.Dispatcher.Invoke((Action)(() =>
            {
                num_bytes = (int.Parse(textBoxSampleNumberOfBytes.Text) > 0) ? int.Parse(textBoxSampleNumberOfBytes.Text) : 4096;
            }));
            send_data_size = (send_data_size < num_bytes) ? send_data_size : num_bytes;


            /* Now send a FEED to the CDC */

            Array.Copy(BitConverter.GetBytes(send_data_size), 0, sendSample, 12, 4);
            sendSample[10] = activity_current;
            sendSample[11] = activity_file;

            event_type = ProductCOMEvents.FeedDone;
            activity_previous = activity_current;
            data_buffers_sent = 0;
            _serialPort.Write(sendSample, 0, Constants.USB_EXTERNAL_BUFFER_SIZE);
        }

        private void buttonActivityPlus_Click(object sender, RoutedEventArgs e)
        {
            int activity = int.Parse(textBoxActivity.Text);
            if (activity < 15)
                ++activity;
            textBoxActivity.Text = (activity).ToString();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            int activity = int.Parse(textBoxActivity.Text);
            if (activity > 1)
                --activity;
            textBoxActivity.Text = (activity).ToString();
        }

        private void buttonListPorts_Click(object sender, RoutedEventArgs e)
        {
            if(comPortsListed)
            {
                SetComPortComboBoxDefaults();
                buttonListPorts.Content = "List Connected";
                comPortsListed = false;
            }
            else
            {
                SetComPortComboBoxListConnected();
                buttonListPorts.Content = "Manual";
                comPortsListed = true;
            }
        }

        private void SetComPortComboBoxListConnected()
        {
            comboBoxPorts.Items.Clear();

            // Get all connected com ports, regardless of their VID/PID
            // TODO: In your application, you should specify a VID/PID like in the comment below
            // List<ComPortProperties> comPortProperties = ComPortManager.GetPortProperties(VID: "0403", PID: "6001");
            List<ComPortProperties> comPortProperties = ComPortManager.GetPortProperties();

            if (comPortProperties.Count == 0)
            {
                MessageBox.Show("No com ports were found");
            }
            else
            {
                foreach (ComPortProperties portProperties in comPortProperties)
                {
                    comboBoxPorts.Items.Add(portProperties);
                }

                // This prevents the combobox selecting 'nothing' after repopulating it
                comboBoxPorts.SelectedIndex = 0;
            }
        }

        // This loads the combo box with ComPortProperties objects so we can have it display a 'friendly name'
        // instead of just 'COM1', and also get the actual com port name after the user has selected it
        private void SetComPortComboBoxDefaults()
        {
            comboBoxPorts.Items.Clear();
            for (int i = 1; i <= 15; i++)
            {
                string name = $"COM{i}";
                comboBoxPorts.Items.Add(new ComPortProperties(
                    ComPortName: name,
                    VID: "",
                    PID: "",
                    FriendlyName: ""
                ));
            }
            comboBoxPorts.SelectedIndex = 0;
        }
    }
}
