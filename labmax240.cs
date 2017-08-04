/*
 * 
 *   Driver Labmax 240
 * 
 * 
 *   por Renato Igleziaz
 *   em  06-28/06/2013
 * 
 */

#region imports
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Ports;
using System.IO;
using System.Windows.Forms;
using System.Data;
using System.ComponentModel;
using System.Threading;
using cDataObject;
using NCalc;
using VirtualMachine40;
#endregion

namespace lab_interfaceamento.Drivers.labmax240
{
    #region Estrutura os parametros de configuração do equipamento
    public class settings
    {
        public int ID = -1;       // id do equipamento
        public int BaudRate = 9600;     // velocidade da porta
        public string Parity = "None";   // Pariedade        
        public int DataBits = 0;        // DataBits
        public string StopBits = "One";    // StopBits        
        public string PortName = "COM0";   // endereço de comunicação da porta serial
    }
    #endregion

    #region rack com as amostras e exames a serem enviados ao equipamento
    public class labmax_Item
    {
        public string codigo_amostra;
        public string amostra_barra;
        public string codigo_exame;
        public string codigo_int;
        public string rack;
        public int    seq;
    }

    public class labmax_ItemsCollections
    {
        private ArrayList vm_items = new ArrayList();

        public void Add(labmax_Item exame)
        {
            // adiciona um item a coleção
            vm_items.Add(exame);
        }

        public labmax_Item ListItem(int index)
        {
            // retorna o item
            return (labmax_Item)vm_items[index];
        }

        public int Count()
        {
            // retorna o total de itens da coleção
            return vm_items.Count;
        }

        public void Clear()
        {
            // limpa toda a coleção
            vm_items.Clear();
        }
    }
    #endregion

    #region Log System
    public class labmax_logInfo
    {
        public string   id;
        public int      resultado;
        public DateTime ocorrencia;
        public string   ocorrencia_descritivo;
        public string   comando_solicitado;
    }

    public class labmax_log
    {
        private ArrayList log_element = new ArrayList();

        public void Add(labmax_logInfo exame)
        {
            // adiciona um item a coleção

            const string stx = "\x02"; // STX = 0x02
            const string etx = "\x03"; // ETX = 0x03
            const string eot = "\x04"; // EOT = 0x04
            const string enq = "\x05"; // ENQ = 0x05
            const string ack = "\x06"; // ACK = 0x06
            const string nak = "\x15"; // NAK = 0x15
            const string etb = "\x17"; // ETB = 0x17

            // se encontrar qualquer um desses caracteres de controle
            // substitui por exibição "humana"
            if (exame.comando_solicitado == stx)
                exame.comando_solicitado = "stx";
            else if (exame.comando_solicitado == etx)
                exame.comando_solicitado = "etx";
            else if (exame.comando_solicitado == eot)
                exame.comando_solicitado = "eot";
            else if (exame.comando_solicitado == enq)
                exame.comando_solicitado = "enq";
            else if (exame.comando_solicitado == ack)
                exame.comando_solicitado = "ack";
            else if (exame.comando_solicitado == nak)
                exame.comando_solicitado = "nak";
            else if (exame.comando_solicitado == etb)
                exame.comando_solicitado = "etb";
            
            log_element.Add(exame);
        }

        public labmax_logInfo ListItem(int index)
        {
            // retorna o item
            return (labmax_logInfo)log_element[index];
        }

        public int Count()
        {
            // retorna o total de itens da coleção
            return log_element.Count;
        }

        public void Clear()
        {
            // limpa toda a coleção
            log_element.Clear();
        }
    }
    #endregion

    #region state of the pins
    public class status
    {
        public bool CD;
        public bool CTS;
        public bool DSR;
        public bool DTR;
        public bool RTS;
    }
    #endregion

    public class labmax240
    {
        #region Ambiente Variaveis
        // internas
        private const string nameofinterface = "Labmax240";
        private cDataBase mdb = new cDataBase();
        private settings setup = new settings();
        private string SQL = "";
        private DataRow dtEquip = null;
        private Thread thread;
        private bool threadCancelOP = false;
        private Form com_form;
        private Label com_bar;
        private ListBox com_listbox;
        private int vm_frame = 0;
        private int vm_delay = 0;
        private string vm_labmaxresponsedata = "";
        private labmax_logInfo loginfo = null;
        private Queue<string> recievedData = new Queue<string>();
        private bool vm_state_thread = false;
        private int vm_qtdgeralresult = -1;
        // externas
        public bool status = false;
        public status vm_status = new status();
        public labmax_log log = new labmax_log();
        public labmax_ItemsCollections ExamesSolicitados;
        public SerialPort porta = new SerialPort();
        
        // Run Function Vars
        private string stringlido = "";
        private string chr_lido = "";
        private string lido = "";
        private bool is_eof = false;
        private int fim_msg = 0;
        private string m_check = "";
        int nr_msg = 0;
        private string new_lido = "";
        int ini_msg = 0;
        int len_msg = 0;
        private string tipo_mensa = "";
        //string send_msg = "";
        private string p_cod_bar = "";
        // header vars
        //string hdr_passwd = "";
        private string hdr_snd_1 = "";
        private string hdr_snd_2 = "";
        private string hdr_rec_1 = "";
        private string hdr_rec_2 = "";
        private string hdr_procid = "";
        private string hdr_version = "";
        private string qry_start = "";
        // paciente
        private string pac_seq = "";
        private string pac_id1 = "";
        private string pac_id2 = "";
        private string pac_id3 = "";
        // order 
        private string ord_seq = "";
        private string ord_spec_id = "";
        private string ord_inst_id = "";
        private string ord_test_id = "";
        private string ord_prior = "";
        //string ord_dt_req = "";
        //string ord_dt_col = "";
        private string ord_action = "";
        private string ord_danger = "";
        private string ord_report = "";
        // result
        private string M_seq = "";
        private string M_exame = "";
        private string M_res_type = "";
        private string M_result = "";
        private string M_unit = "";
        private string M_ref = "";
        private string M_flag = "";
        //string M_nature = "";
        private string M_status = "";
        private string M_data = "";
        //string M_instru = "";
        //string M_hora = "";
        // comment
        private string com_text = "";
        // Q inquiry
        private string qry_seq = "";
        //string qry_car = "";
        //string qry_pos = "";
        //string qry_dct = "";
        //string qry_ct = "";
        //string qry_end = "";
        private string qry_test_id = "";
        private string qry_status = "";
        //  terminator message
        private string term_code = "";
        private string term_name = "";
        
        #endregion

        #region Controle de Dados
        public const string stx = "\x02"; // STX = 0x02
        public const string etx = "\x03"; // ETX = 0x03
        public const string eot = "\x04"; // EOT = 0x04
        public const string enq = "\x05"; // ENQ = 0x05
        public const string ack = "\x06"; // ACK = 0x06
        public const string _lf = "\x0A"; //  LF = 0x0A
        public const string _cr = "\x0D"; //  CR = 0x0D
        public const string nak = "\x15"; // NAK = 0x15
        public const string etb = "\x17"; // ETB = 0x17
        #endregion

        #region Delimitadores
        public const string fs = @"|"; // Field     delimiter
        public const string rd = @"\"; // Repeat    delimiter
        public const string cd = @"^"; // Component delimiter
        public const string ed = @"&"; // Escape    delimiter
        #endregion

        #region Inicializa
        public labmax240(string id_equipamento)
        {
            // banco de dados
            mdb.strProvider = Global.ProviderDb;
            mdb.AppPath = Global.PathApp;

            // carrega ambiente
            SQL = "SELECT * FROM EQUIPAMENTOS WHERE CODIGO=" + id_equipamento;
            cReturnDataTable tmpEquip = mdb.Abrir_DataTable(SQL);
            if (!tmpEquip.Status)
            {
                // retorna erro
                return;
            }

            if (tmpEquip.RecordCount == 0)
            {
                // retorna erro
                return;
            }

            try
            {
                dtEquip = tmpEquip.dt.Rows[0];
                setup.ID = int.Parse(id_equipamento);
                setup.PortName = dtEquip["PORTA"].ToString();
                setup.BaudRate = int.Parse(dtEquip["VELOCIDADE"].ToString());
                setup.Parity = dtEquip["PARIEDADE"].ToString();
                setup.DataBits = int.Parse(dtEquip["DATABITS"].ToString());
                setup.StopBits = dtEquip["STOPBITS"].ToString();
            }
            catch
            {
                // não conseguiu carregar o setup de configuração
                return;
            }

            // OK
            status = true;
        }
        #endregion

        #region Retornar o conteudo de um determinado campo e componente
        public string seekfield(string p_string, int p_field, int p_component = 1)
        {
            int inicio = 0;
            int tamanho = 0;
            string p_seekfield = "";

            if (p_field == 1)
                inicio = 0;
            else
                inicio = At(p_string, fs, p_field-1)+1;

            tamanho = At(p_string, fs, p_field) - inicio;
            if (tamanho < 0)
                tamanho = p_string.Length - inicio;

            p_seekfield = Substring(p_string, inicio, tamanho).Trim();

            if (p_component == 1)
                inicio = 0;
            else
                inicio = At(p_seekfield, cd, p_component-1)+1;

            tamanho = At(p_seekfield, cd, p_component) - inicio;
            if (tamanho <= 0)
                tamanho = p_seekfield.Length - inicio;

            p_seekfield = Substring(p_seekfield, inicio, tamanho).Trim();

            return p_seekfield;
        }

        public int At(string input, string search, int ocorrence = 1)
        {
            // mesma função que a string.Index()
            // mas tive que recriar a função por
            // causa do parametro ocorrencia

            int intern_ocorrence = 1;

            for (int x = 0; x < input.Length; x++)
            {
                if (input.Substring(x, 1) == search)
                {
                    if (ocorrence == intern_ocorrence)
                        return x;
                    else
                        intern_ocorrence++;
                }
            }

            return -1;
        }

        public string Substring(string input, int start, int size)
        {
            // Substring
            string retorno = "";
            int pos = start;
            int tam = 0;

            if (input.Trim() == "")
                return retorno;

            for (; ; )
            {
                if (tam >= size)
                    break;

                retorno += input.Substring(pos, 1);

                pos++;
                tam++;
            }

            return retorno;
        }
        #endregion

        #region Retornar o código hexadecimal do número
        public string GetInt2Hexa(int input)
        {
            string hex = "";
            const string str_base = "0123456789ABCDEF";
            int p_number = input;

            for (; ; )
            {
                hex = str_base.Substring((p_number % 16), 1) + hex;

                p_number = (int)(p_number / 16);
                if (p_number == 0)
                    break;
            }
            
            return hex;
        }
        #endregion

        #region Build CheckSum
        public string GetCheckSum(string p_string)
        {
            int i_checksum = 0;
            int i_char = 0;
            string formatresult = "";

            for (int i = 0; i < p_string.Length; i++)
            {
                i_char = GetAsc(p_string.Substring(i, 1));
                i_checksum = (i_checksum + i_char);
                i_checksum = i_checksum % 256;
            }

            formatresult = "00" + GetInt2Hexa(i_checksum);

            return formatresult.Substring(formatresult.Length - 2, 2);
        }
        #endregion

        #region Asc em C#
        public int GetAsc(string _char)
        {
            char c = char.Parse(_char);
            return (int)c;
        }
        #endregion

        #region Porta de Comunicação
        public bool OpenPort(settings optional = null, bool EventHandler = false)
        {
            // vars
            bool error = false;

            if (!status)
                return false;

            // se a porta estiver aberta
            if (porta.IsOpen) 
                porta.Close();

            // flags de conexão padrão
            settings now = null;

            if (optional != null)
                now = optional;
            else
                now = setup;

            // velocidade da porta
            porta.BaudRate = now.BaudRate;
            // Pariedade
            porta.Parity = (Parity)Enum.Parse(typeof(Parity), now.Parity);
            // DataBits
            porta.DataBits = now.DataBits;
            // StopBits
            porta.StopBits = (StopBits)Enum.Parse(typeof(StopBits), now.StopBits);
            // endereço de comunicação da porta serial
            porta.PortName = now.PortName;
            
            try
            {
                porta.Open();
            }
            catch (UnauthorizedAccessException) 
            {
                // log
                loginfo = new labmax_logInfo();
                loginfo.id = "OpenPort()";
                loginfo.ocorrencia = DateTime.Now;
                loginfo.ocorrencia_descritivo = "UnauthorizedAccessException";
                loginfo.resultado = -1;
                log.Add(loginfo);

                error = true; 
            }
            catch (IOException) 
            {
                // log
                loginfo = new labmax_logInfo();
                loginfo.id = "OpenPort()";
                loginfo.ocorrencia = DateTime.Now;
                loginfo.ocorrencia_descritivo = "IOException";
                loginfo.resultado = -1;
                log.Add(loginfo);

                error = true; 
            }
            catch (ArgumentException) 
            {
                // log
                loginfo = new labmax_logInfo();
                loginfo.id = "OpenPort()";
                loginfo.ocorrencia = DateTime.Now;
                loginfo.ocorrencia_descritivo = "ArgumentException";
                loginfo.resultado = -1;
                log.Add(loginfo);

                error = true; 
            }

            if (error)
            {
                // log
                loginfo = new labmax_logInfo();
                loginfo.id = "OpenPort()";
                loginfo.ocorrencia = DateTime.Now;
                loginfo.ocorrencia_descritivo = "Não pode abrir a porta de comunicação.";
                loginfo.resultado = -1;
                log.Add(loginfo);
            }
            else
            {
                // registra primeiro pin state
                vm_status.CD  = porta.CDHolding;
                vm_status.CTS = porta.CtsHolding;
                vm_status.DSR = porta.DsrHolding;
                vm_status.DTR = porta.DtrEnable;
                vm_status.RTS = porta.RtsEnable;
            }

            if (!porta.IsOpen)
                return false;

            if (error)
                return false;

            // comm event handler
            porta.DataReceived += new SerialDataReceivedEventHandler(port_DataReceived);

            //if (EventHandler)
            //{
            //    //porta.PinChanged += new SerialPinChangedEventHandler(port_PinChanged);
            //}

            return true;
        }

        public void ClosePort()
        {
            // fecha porta de comunicação

            try
            {
                if (porta.IsOpen)
                    porta.Close();
            }
            catch { }
        }

        public int SendCommandToPort(string mensagem, bool usageETB = false)
        {
            // envia um comando para uma porta de comunicação

            // lista de retornos
            // 0 - retorno OK
            // 1 - falha na escrita da porta
            // 2 - mensagem não confirmada
            // 3 - esgotou tentativas, por favor, tentar novamente em alguns segundos.
            // 4 - esgotou tempo de resposta
            // 5 - contenção, aguardar antes de transmitir
            // 6 - a porta não estava aberta
            // 7 - a mensagem a ser enviada estava vazia
            
            // se a porta não estiver aberta
            if (!porta.IsOpen)
            {
                // log
                loginfo = new labmax_logInfo();
                loginfo.id = "SendCommandToPort()";
                loginfo.ocorrencia = DateTime.Now;
                loginfo.ocorrencia_descritivo = "a porta não estava aberta";
                loginfo.resultado = 6;
                loginfo.comando_solicitado = mensagem;
                log.Add(loginfo);
                return 6;
            }
            
            if (mensagem.Trim() == "")
            {
                loginfo = new labmax_logInfo();
                loginfo.id = "SendCommandToPort()";
                loginfo.ocorrencia = DateTime.Now;
                loginfo.ocorrencia_descritivo = "a mensagem a ser enviada estava vazia";
                loginfo.resultado = 7;
                loginfo.comando_solicitado = mensagem;
                log.Add(loginfo);
                return 7;
            }

            string new_msg = "";
            int tentativas = 6;
            int timer = 15;

            // não é um caracter de controle
            if (mensagem.Length > 1)
            {
                // calcula o frame
                vm_frame = (vm_frame + 1) % 8;     
          
                // estrutura mensagem
                // <STX>[FN][TEXT]<ETB>[C1][C2]<CR><LF>
                new_msg = vm_frame.ToString().Substring(0, 1) + mensagem;
                
                if (!usageETB)
                    new_msg = new_msg + etx;
                else
                    new_msg = new_msg + etb;

                new_msg = stx + new_msg + GetCheckSum(new_msg) + Environment.NewLine;
                mensagem = new_msg;
            }

            for (; ; )
            {
                if (tentativas == 0)
                    break;

                // reset flags
                vm_labmaxresponsedata = "";

                // dorme o serviço caso precise
                if (vm_delay > 0)
                    System.Threading.Thread.Sleep(vm_delay);

                // escreve mensagem na porta
                try
                {
                    for (int i = 0; i < mensagem.Length; i++)
                    {
                        porta.Write(mensagem.Substring(i, 1));
                    }
                }
                catch
                {
                    // erro de escrita na porta
                    loginfo = new labmax_logInfo();
                    loginfo.id = "SendCommandToPort()";
                    loginfo.ocorrencia = DateTime.Now;
                    loginfo.ocorrencia_descritivo = "falha na escrita da porta";
                    loginfo.resultado = 1;
                    loginfo.comando_solicitado = mensagem;
                    log.Add(loginfo);
                    return 1;
                }

                // Aguarda retorno caso a mensagem enviada nao seja ACK ou EOT
                if ((mensagem != ack) && (mensagem != eot))
                {
                    // aguarda o retorno
                    for (; ; )
                    {
                        if (timer == 0)
                            break;

                        try
                        {
                            vm_labmaxresponsedata = this.ReadBuffer();
                        }
                        catch 
                        { 
                            vm_labmaxresponsedata = ""; 
                        }

                        if (vm_labmaxresponsedata.Trim() != "")
                            break;

                        // dorme por 1s
                        System.Threading.Thread.Sleep(1000);
                        timer--;
                    }

                    if (vm_labmaxresponsedata.Trim() == "" && mensagem == enq)
                    {
                        // saida por timeout
                        try
                        {
                            porta.Write(eot);
                        }
                        catch
                        {
                            // erro de escrita na porta
                            loginfo = new labmax_logInfo();
                            loginfo.id = "SendCommandToPort()";
                            loginfo.ocorrencia = DateTime.Now;
                            loginfo.ocorrencia_descritivo = "falha na escrita da porta";
                            loginfo.resultado = 1;
                            loginfo.comando_solicitado = mensagem;
                            log.Add(loginfo);
                            return 1;
                        }

                        // saida por timeout
                        loginfo = new labmax_logInfo();
                        loginfo.id = "SendCommandToPort()";
                        loginfo.ocorrencia = DateTime.Now;
                        loginfo.ocorrencia_descritivo = "esgotou tempo de resposta";
                        loginfo.resultado = 4;
                        loginfo.comando_solicitado = mensagem;
                        log.Add(loginfo);
                        return 4;
                    }
                    else if ((vm_labmaxresponsedata == enq || vm_labmaxresponsedata == nak) && mensagem == enq)
                    {
                        // saida por contenção
                        
                        // aguarda 1s antes de enviar uma nova solicitação
                        System.Threading.Thread.Sleep(1000);
                    }
                    else if (vm_labmaxresponsedata == ack)
                    {
                        // recebeu um ACK

                        // retorno OK
                        loginfo = new labmax_logInfo();
                        loginfo.id = "SendCommandToPort()";
                        loginfo.ocorrencia = DateTime.Now;
                        loginfo.ocorrencia_descritivo = "ACK - Comunicação OK";
                        loginfo.resultado = 0;
                        loginfo.comando_solicitado = mensagem;
                        log.Add(loginfo);
                        return 0;
                    }
                    else if (vm_labmaxresponsedata == nak)
                    {
                        // mensagem não confirmada

                        // aguarda os 10s antes de re-escrever
                        System.Threading.Thread.Sleep(10000);
                    }
                    tentativas--;
                }
                else
                    tentativas = 0;

                // reconfigura timer de espera
                timer = 15;
            }

            // controle de saida
            if (tentativas == 0 && mensagem != ack && mensagem != eot && mensagem != enq)
            {
                // Mensagem nao confirmada nas tentativas efetuadas
                try
                {
                    porta.Write(eot);
                }
                catch { };

                loginfo = new labmax_logInfo();
                loginfo.id = "SendCommandToPort()";
                loginfo.ocorrencia = DateTime.Now;
                loginfo.ocorrencia_descritivo = "mensagem não confirmada";
                loginfo.resultado = 2;
                loginfo.comando_solicitado = mensagem;
                log.Add(loginfo);
                return 2;
            }
            else if (tentativas == 0 && mensagem == enq)
            {
                // as 6 tentativas foram usadas
                try
                {
                    porta.Write(eot);
                }
                catch { };

                loginfo = new labmax_logInfo();
                loginfo.id = "SendCommandToPort()";
                loginfo.ocorrencia = DateTime.Now;
                loginfo.ocorrencia_descritivo = "esgotou tentativas, por favor, tentar novamente em alguns segundos.";
                loginfo.resultado = 3;
                loginfo.comando_solicitado = mensagem;
                log.Add(loginfo);
                return 3;
            }

            // retorno OK
            loginfo = new labmax_logInfo();
            loginfo.id = "SendCommandToPort()";
            loginfo.ocorrencia = DateTime.Now;
            loginfo.ocorrencia_descritivo = "ACK - Comunicação OK";
            loginfo.resultado = 0;
            loginfo.comando_solicitado = mensagem;
            log.Add(loginfo);
            return 0;
        }

        private void port_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            // controla eventos de resposta da porta de comunicação
           
            try
            {
                recievedData.Enqueue(porta.ReadExisting());
            }
            catch { }
        }

        public string ReadBuffer()
        {
            // le o proximo buffer livre

            return (recievedData.Count > 0) ? recievedData.Dequeue() : "";
        }

        private void port_PinChanged(object sender, SerialPinChangedEventArgs e)
        {
            // exibe os status dos pinos

            
        }
        #endregion

        #region Multi-Thread / Fluxo de transmissão de dados

        public bool StartCOMM(Form                    thisForm, 
                              Label                   thisBar        = null, 
                              ListBox                 thisListBox    = null,
                              labmax_ItemsCollections thisColecao    = null,
                              int                     qtdgeralresult = -1
                             ) 
        {
            // inicia os processos de transmissão de dados
            // por meio de multi-processamento

            threadCancelOP = false;
            com_form = thisForm;
            com_bar = thisBar;
            com_listbox = thisListBox;
            vm_qtdgeralresult = qtdgeralresult;

            if (thisColecao == null)
            {
                MessageBox.Show("A coleção não está disponível!", "StartCOMM");
                return false;
            }
            else
                ExamesSolicitados = thisColecao;

            try
            {
                thread = new Thread(new ThreadStart(run));
                thread.SetApartmentState(ApartmentState.STA);
                thread.IsBackground = true;
                thread.Priority = ThreadPriority.Normal;
                thread.Start();
            }
            catch
            {
                return false;
            }

            return true;
        }

        public bool CancelCOMM()
        {
            // para um serviço de transmissão

            try
            {
                threadCancelOP = true;
            }
            catch
            {
                return false;
            }

            return true;
        }

        private static void InvokeIfRequired(System.Windows.Forms.Control c, Action<System.Windows.Forms.Control> action)
        {
            if (c.InvokeRequired)
            {
                c.Invoke(new Action(() => action(c)));
            }
            else
            {
                action(c);
            }
        }

        private void AddText(string mensagem)
        {
            InvokeIfRequired(com_listbox, x =>
            {
                com_listbox.Items.Add(mensagem);

                if (com_listbox.Items.Count > 0)
                {
                    com_listbox.SelectedIndex = (com_listbox.Items.Count - 1);
                }
            });                
        }

        private string get_codigo_amostra(string codigo_barra)
        {
            // separa a string
            string setor = "";
            string codigoposto = "";
            string amostrareal = "";
            cReturnDataTable tmprec;

            try
            {
                setor = codigo_barra.Substring(0, 2);
                codigoposto = codigo_barra.Substring(2, 2);
                amostrareal = codigo_barra.Substring(4, codigo_barra.Length - 4);
                amostrareal = int.Parse(amostrareal).ToString();
            }
            catch
            {
                setor = "";
                codigoposto = "";
                amostrareal = "";
            }

            if (setor.Trim() == "" || codigoposto.Trim() == "" || amostrareal.Trim() == "")
            {
                return "";
            }

            // pesquisa os dados da amostra
            SQL = "SELECT ";
            SQL += "AMOSTRAS.CODIGO_AMOSTRA ";
            SQL += "FROM AMOSTRAS ";
            SQL += "WHERE AMOSTRAS.CODIGO_POSTOCOLETA=" + int.Parse(codigoposto).ToString() + " ";
            SQL += "AND AMOSTRAS.AMOSTRA='" + amostrareal + "' ";
            tmprec = mdb.Abrir_DataTable(SQL);
            if (!tmprec.Status)
                return "";

            if (tmprec.RecordCount == 0)
            {
                return "";
            }

            DataRow dt = tmprec.dt.Rows[0];
            return dt["CODIGO_AMOSTRA"].ToString();
        }

        private void run() 
        {
            // serviço de transmissão de dados

            // abre a porta do equipamento
            if (!this.OpenPort(null, true))
            {
                InvokeIfRequired(com_bar, X => { com_bar.Text = "Erro abertura de porta"; });
                return;
            }

            // marca que a thread iniciou corretamente
            vm_state_thread = true;
            // controle de legenda
            bool umavez = false;
            cUpdate update = new cUpdate();
            string codigo_amostra = "";
            int controla_result_final = 0;
            cReturnDataTable tmprec;
            string codigo_exame = "";
            string flag = "";
            cInsertInto inc = new cInsertInto();

            #region Reset Vars
            stringlido = "";
            chr_lido = "";
            lido = "";
            is_eof = false;
            fim_msg = 0;
            m_check = "";
            new_lido = "";
            tipo_mensa = "";
            p_cod_bar = "";
            // header vars
            hdr_snd_1 = "Prestige24i";
            hdr_snd_2 = "SYSTEM1";
            hdr_rec_1 = "";
            hdr_rec_2 = "";
            hdr_procid = "";
            hdr_version = "";
            qry_start = "";
            // paciente
            pac_seq = "";
            pac_id1 = "";
            pac_id2 = "";
            pac_id3 = "";
            // order 
            ord_seq = "";
            ord_spec_id = "";
            ord_inst_id = "";
            ord_test_id = "";
            ord_prior = "";
            ord_action = "";
            ord_danger = "";
            ord_report = "";
            // result
            M_seq = "";
            M_exame = "";
            M_res_type = "";
            M_result = "";
            M_unit = "";
            M_ref = "";
            M_flag = "";
            M_status = "";
            M_data = "";
            // comment
            com_text = "";
            // Q inquiry
            qry_seq = "";
            qry_test_id = "";
            qry_status = "";
            //  terminator message
            term_code = "";
            term_name = "";
            #endregion

            // laço principal de leitura
            for (; ; )
            {
                // cancela a operação caso seja solicitado                
                if (threadCancelOP)
                    break;

                // verifica a resposta do equipamento 
                stringlido = this.ReadBuffer();

                if (stringlido.Trim() != "")
                {
                    // somente em caso de depuração
                    //AddText(stringlido);

                    if (!umavez)
                    {
                        InvokeIfRequired(com_bar, X => { com_bar.Text = "Analisando Buffer..."; });
                        umavez = true;
                    }
                }

                for (int pos = 0; pos < stringlido.Length; pos++)
                {
                    // le caracter a caracter
                    chr_lido = stringlido.Substring(pos, 1);

                    #region (case) interpreta o que foi recebido
                    switch (chr_lido)
                    {
                        case "":
                            {
                                System.Threading.Thread.Sleep(500);
                                break;
                            }
                        case stx:
                            {
                                lido = "";
                                is_eof = false;
                                break;
                            }
                        case enq:
                            {
                                if (vm_delay > 0)
                                {
                                    System.Threading.Thread.Sleep(vm_delay);
                                }

                                try
                                {
                                    if (porta.IsOpen)
                                        porta.Write(ack);
                                }
                                catch { }

                                lido = "";

                                break;
                            }
                        case eot:
                            {
                                is_eof = true;
                                chr_lido = "";
                                break;
                            }
                    } // end switch
                    #endregion

                    // acumula a linha de mensagem
                    lido = lido + chr_lido;

                    #region Finaliza a Mensagem
                    if (chr_lido == _lf) // final de linha
                    {
                        // ALERT: [Recebeu] + lido
                        AddText("[Recebeu] " + lido);

                        if (lido.Substring(0, 1) != stx)
                        {
                            //ALERT: mensagem quebrada
                            AddText("Mensagem Quebrada");

                            if (vm_delay > 0)
                                System.Threading.Thread.Sleep(vm_delay);

                            try
                            {
                                if (porta.IsOpen)
                                    porta.Write(nak);
                            }
                            catch { }

                            // ALERT: NAK mensagem quebrada
                            AddText("NAK mensagem quebrada");

                            lido = "";
                            goto VemNanar;
                        }

                        if (this.At(lido, etx) != 0)
                        {
                            // end of text
                            fim_msg = this.At(lido, etx);
                        }
                        else
                        {
                            // end of Block
                            fim_msg = this.At(lido, etb);
                        }

                        // recupera o checksum
                        m_check = this.Substring(lido, fim_msg + 1, 2);
                        lido = this.Substring(lido, 1, lido.Length - 5);

                        // calcula o checksum
                        if (m_check == this.GetCheckSum(lido))
                        {
                            if (vm_delay > 0)
                                System.Threading.Thread.Sleep(vm_delay);

                            // ALERT: [Envia] ACK
                            AddText("[Envia] ACK");

                            try
                            {
                                if (porta.IsOpen)
                                    porta.Write(ack);
                            }
                            catch { }
                        }
                        else
                        {
                            if (vm_delay > 0)
                                System.Threading.Thread.Sleep(vm_delay);

                            // ALERT: [Envia] NAK CHECKSUM
                            AddText("[Envia] NAK CHECKSUM");

                            try
                            {
                                if (porta.IsOpen)
                                    porta.Write(nak);
                            }
                            catch { }

                            lido = "";
                            goto VemNanar;
                        }

                        // estado finito - verificação do tipo de mensagem
                        nr_msg = 0;

                        for (int x = 0; x < lido.Length; x++)
                            if (this.Substring(lido, x, 1) == _cr)
                                nr_msg++;

                        new_lido = lido;

                        // laço para cada mensagem constante na string recebida
                        for (int x = 1; x <= nr_msg; x++)
                        {
                            if (x == 1)
                                ini_msg = 1;
                            else
                                ini_msg = this.At(new_lido, _cr, x - 1) + 1;

                            if (x == nr_msg)
                                len_msg = new_lido.Length - ini_msg;
                            else
                                len_msg = this.At(new_lido, _cr, x) - ini_msg + 1;

                            lido = this.Substring(new_lido, ini_msg, len_msg);

                            tipo_mensa = this.seekfield(lido, 1);
                            //send_msg = "";
                            //vm_frame = int.Parse(this.Substring(lido, 2, 1));

                            #region verificação do tipo de mensagem

                            switch (tipo_mensa)
                            {
                                case "H":
                                    {
                                        // Message Header Record
                                        //hdr_passwd = seekfield(lido, 4);    // Access Password          
                                        hdr_snd_1 = seekfield(lido, 5, 1);  // Sender Name or ID    
                                        hdr_snd_2 = seekfield(lido, 5, 2);
                                        hdr_rec_1 = seekfield(lido, 10, 1); // Receiver ID 
                                        hdr_rec_2 = seekfield(lido, 10, 2); // Receiver ID
                                        hdr_procid = seekfield(lido, 12);   // Processing ID
                                        hdr_version = seekfield(lido, 13);  // Version No.
                                        qry_start = "";                     // Zera query
                                        AddText("Recebido (H) Message");
                                        break;
                                    }
                                case "P":
                                    {
                                        // antes de checar o novo paciente
                                        // verifica se jah teve alguma outra
                                        // amostra antes, para processar os 
                                        // cálculos.

                                        if (p_cod_bar.Trim().Length > 0)
                                        {
                                            // processa cálculos
                                            ProccessCalculate(p_cod_bar);
                                        }

                                        // Patient Information Record
                                        pac_seq = seekfield(lido, 2); // Sequence number
                                        pac_id1 = seekfield(lido, 3); // Patient ID
                                        pac_id2 = seekfield(lido, 4); // Laboratory Assigned Patient ID (codigo de barras)
                                        pac_id3 = seekfield(lido, 5); // Patient ID n3 
                                        p_cod_bar = pac_id2;
                                        AddText("Recebido Paciente '" + pac_seq + "'");
                                        break;
                                    }
                                case "O":
                                    {
                                        // Measurement Order Record
                                        ord_seq = seekfield(lido, 2);       // Sequence Number
                                        ord_spec_id = seekfield(lido, 3);   // Specimen ID
                                        ord_inst_id = seekfield(lido, 4);   // Instrument Specimen ID
                                        ord_test_id = seekfield(lido, 5);   // Universal Test ID
                                        ord_prior = seekfield(lido, 6);     // Priority
                                        //ord_dt_req = seekfield(lido, 7);    // Requested/Ordered Date and Time
                                        //ord_dt_col = seekfield(lido, 8);    // Specimen Collection Daten and time
                                        ord_action = seekfield(lido, 12);   // Action Code
                                        ord_danger = seekfield(lido, 13);   // Danger Code
                                        ord_report = seekfield(lido, 26);   // Report Types 
                                        p_cod_bar = ord_spec_id;
                                        break;
                                    }
                                case "R":
                                    {
                                        // Measurement Result Record
                                        M_seq = seekfield(lido, 2);             // Sequence Number
                                        M_exame = seekfield(lido, 3, 4);        // Universal Test ID
                                        M_res_type = seekfield(lido, 3, 8);     // Tipo de resultado (DOSE, COFF, RLU)
                                        M_result = seekfield(lido, 4);          // Data or Measurement Value
                                        M_unit = seekfield(lido, 5);            // Units
                                        M_ref = seekfield(lido, 6);             // Reference Range
                                        M_flag = seekfield(lido, 7);            // Result Abnormal Flags
                                        //M_nature = seekfield(lido, 8);          // Nature of Abnormality Testing
                                        M_status = seekfield(lido, 9);          // Result Status
                                        M_data = seekfield(lido, 13);           // Date/Time Test Completed
                                        //M_instru = seekfield(lido, 14);         // Instrument Identification
                                        //M_data = this.Substring(M_data, 0, 8);  // Separa data da hora
                                        //M_hora = this.Substring(M_data, 8, 6);  // Separa hora da data

                                        switch (M_flag.Trim())
                                        {
                                            case "L":
                                                {
                                                    M_flag = "L-Lower than the lower limit";
                                                    break;
                                                }
                                            case "H":
                                                {
                                                    M_flag = "H-Higher than the upper limit";
                                                    break;
                                                }
                                            case "LL":
                                                {
                                                    M_flag = "LL-Lower than the lowest limit";
                                                    break;
                                                }
                                            case "HH":
                                                {
                                                    M_flag = "HH-Higher than the highest limit";
                                                    break;
                                                }
                                            case "<":
                                                {
                                                    M_flag = "<-Lower than the low limit absolute value";
                                                    break;
                                                }
                                            case ">":
                                                {
                                                    M_flag = ">-Higher than the high limit absolute value";
                                                    break;
                                                }
                                            case "N":
                                                {
                                                    M_flag = "N-Normal";
                                                    break;
                                                }
                                            case "A":
                                                {
                                                    M_flag = "A-Abnormal";
                                                    break;
                                                }
                                            case "U":
                                                {
                                                    M_flag = "U-the significant digit goes up";
                                                    break;
                                                }
                                            case "D":
                                                {
                                                    M_flag = "D-the significant digit goes down";
                                                    break;
                                                }
                                            case "B":
                                                {
                                                    M_flag = "B-good";
                                                    break;
                                                }
                                            case "W":
                                                {
                                                    M_flag = "W-wrong";
                                                    break;
                                                }
                                        }

                                        codigo_amostra = get_codigo_amostra(p_cod_bar);
                                        if (codigo_amostra != "")
                                        {
                                            InvokeIfRequired(com_bar, X => { com_bar.Text = "Armazenando a ordem '" + p_cod_bar + "'..."; });

                                            // aqui vai gravar o resultado na base de dados
                                            update.Clear();
                                            update.Add("R_RESULT", M_result, true);
                                            update.Add("R_UNIT", M_unit, true);
                                            update.Add("R_REF", M_ref, true);
                                            update.Add("R_FLAG", M_flag, true);
                                            update.Add("R_STATUS", M_status, true);
                                            update.Add("R_DATA", DateTime.Now.ToString("MM/dd/yyyy HH:mm"), true);

                                            // filtra item para gravar o resultado
                                            SQL = "WHERE CODIGO_EQUIPA=" + setup.ID.ToString() + " ";
                                            SQL += "AND CODIGO_AMOSTRA='" + codigo_amostra + "' ";
                                            SQL += "AND R_UNIVERSALID='" + M_exame + "' ";
                                            //SQL += "AND R_ORDER=" + ord_seq + " ";
                                            //SQL += "AND R_SEQ=" + M_seq + " ";

                                            SQL = update.Update("INTERFACE_CAMPOS", SQL);
                                            mdb.SQLExecute(SQL);

                                            controla_result_final++;

                                            // notifica o status do exame
                                            SQL = "SELECT ";
                                            SQL += "INTERFACE_CAMPOS.CODIGO_AMOSTRA, ";
                                            SQL += "AMOSTRAS_EXAMES.SEQUENCIA, ";
                                            SQL += "INTERFACE_CAMPOS.CODIGO_EXAME, ";
                                            SQL += "INTERFACE_CAMPOS.R_FLAG ";
                                            SQL += "FROM INTERFACE_CAMPOS ";
                                            SQL += "INNER JOIN AMOSTRAS ON ";
                                            SQL += "AMOSTRAS.CODIGO_AMOSTRA=INTERFACE_CAMPOS.CODIGO_AMOSTRA ";
                                            SQL += "INNER JOIN AMOSTRAS_EXAMES ON ";
                                            SQL += "AMOSTRAS_EXAMES.CODIGO_AMOSTRA=AMOSTRAS.CODIGO_AMOSTRA ";
                                            SQL += "AND AMOSTRAS_EXAMES.CODIGO_EXAME=INTERFACE_CAMPOS.CODIGO_EXAME ";
                                            SQL += "WHERE INTERFACE_CAMPOS.CODIGO_AMOSTRA='" + codigo_amostra + "' ";
                                            SQL += "AND INTERFACE_CAMPOS.CODIGO_EQUIPA=" + setup.ID.ToString() + " ";
                                            tmprec = mdb.Abrir_DataTable(SQL);
                                            if (!tmprec.Status)
                                                goto GoOut;

                                            codigo_exame = "";
                                            flag = "";

                                            foreach (DataRow dtgravastatus in tmprec.dt.Rows)
                                            {
                                                if (codigo_exame != dtgravastatus["CODIGO_EXAME"].ToString())
                                                {
                                                    // novo exame
                                                    codigo_exame = dtgravastatus["CODIGO_EXAME"].ToString();
                                                    flag = dtgravastatus["SEQUENCIA"].ToString();

                                                    // grava inicialmente o status normal
                                                    SQL = "WHERE CODIGO_AMOSTRA='" + codigo_amostra + "' ";
                                                    SQL += "AND EQUIPAMENTO=" + setup.ID.ToString() + " ";
                                                    SQL += "AND SEQUENCIA=" + flag + " ";

                                                    inc.Clear();
                                                    inc.Add("STATUS", "Normal", true);
                                                    SQL = inc.Update("INTERFACE_BOT", SQL);
                                                    if (!mdb.SQLExecute(SQL))
                                                        return;
                                                }

                                                if (dtgravastatus["R_FLAG"].ToString().Trim().Length > 0)
                                                {
                                                    if (dtgravastatus["R_FLAG"].ToString().Substring(0, 1) != "N" && dtgravastatus["R_FLAG"].ToString().Substring(0, 1) != "B")
                                                    {
                                                        if (dtgravastatus["R_FLAG"].ToString().Substring(0, 1) == "A")
                                                        {
                                                            // absurdo
                                                            SQL = "WHERE CODIGO_AMOSTRA='" + codigo_amostra + "' ";
                                                            SQL += "AND EQUIPAMENTO=" + setup.ID.ToString() + " ";
                                                            SQL += "AND SEQUENCIA=" + flag + " ";

                                                            inc.Clear();
                                                            inc.Add("STATUS", "Absurdo", true);
                                                            SQL = inc.Update("INTERFACE_BOT", SQL);
                                                            if (!mdb.SQLExecute(SQL))
                                                                return;
                                                        }
                                                        else 
                                                        {
                                                            // alterado
                                                            SQL = "WHERE CODIGO_AMOSTRA='" + codigo_amostra + "' ";
                                                            SQL += "AND EQUIPAMENTO=" + setup.ID.ToString() + " ";
                                                            SQL += "AND SEQUENCIA=" + flag + " ";

                                                            inc.Clear();
                                                            inc.Add("STATUS", "Alterado", true);
                                                            SQL = inc.Update("INTERFACE_BOT", SQL);
                                                            if (!mdb.SQLExecute(SQL))
                                                                return;
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        AddText("Resultado OK: " + p_cod_bar + "^" + M_exame + "^" + M_result + "^" + M_unit + "^" + M_flag);
                                        break;
                                    }
                                case "C":
                                    {
                                        // Comment Record
                                        com_text = seekfield(lido, 4, 2);  // Comment

                                        // ALERT: "COMMENT Text:" + com_text
                                        AddText("COMMENT Text: " + com_text);
                                        break;
                                    }
                                case "Q":
                                    {
                                        // Enquiry Record
                                        qry_seq = seekfield(lido, 2);       // Sequence Number
                                        qry_start = seekfield(lido, 3);     // Starting Range ID Number 
                                        qry_test_id = seekfield(lido, 5);   // Test item ID
                                        qry_status = seekfield(lido, 13);   // Condition code

                                        // ALERT: "RX QUERY " + qry_start
                                        AddText("A maquina solicitou exames...");
                                        break;
                                    }
                                case "L":
                                    {
                                        // Message Terminator Record
                                        term_code = seekfield(lido, 3);
                                        term_code = term_code.Substring(0, 1);
                                        term_name = "NO MESSAGE";

                                        if (term_code == "N")
                                            term_name = "NORMAL";
                                        else if (term_code == "T")
                                            term_name = "SENDER ABORTED";
                                        else if (term_code == "R")
                                            term_name = "RECEIVER REQUESTED ABORT";
                                        else if (term_code == "E")
                                            term_name = "UNKNOWN SYSTEM ERROR";
                                        else if (term_code == "Q")
                                            term_name = "ERROR IN LAST REQUEST FOR INFORMATION";
                                        else if (term_code == "I")
                                            term_name = "NO INFORMATION AVAILABLE FROM LAST QUERY";
                                        else if (term_code == "F")
                                            term_name = "FINAL REQUEST FOR INFORMATION PROCESSED";

                                        // ALERT: " RX TERMINATOR: " + term_name
                                        AddText("L TERMINATOR: " + term_code + "-" + term_name);
                                        break;
                                    }

                            } // end - verificação de tipo de mensagem

                            #endregion

                            lido = "";
                        }
                    }
                    #endregion

                    if (qry_start.Trim() != "" && is_eof)
                    {
                        // envia amostra para o equipamento
                        queryresp();

                        // reset
                        qry_start = "";
                        is_eof = false;
                        umavez = false;
                        InvokeIfRequired(com_bar, X => { com_bar.Text = "Aguardando Resultado..."; });
                    }
                    else if (is_eof && controla_result_final >= vm_qtdgeralresult)
                    {
                        // verifica que todas as solicitações foram respondidas
                        // pelo equipamento

                        // processa cálculos
                        ProccessCalculate(p_cod_bar);
                        
                        InvokeIfRequired(com_bar, X => { com_bar.Text = "Concluído com sucesso..."; });
                        System.Threading.Thread.Sleep(3000);
                        goto GoOut;
                    }

                // dorme ...
                VemNanar: ;
                    System.Threading.Thread.Sleep(10);

                } // le o proximo caracter                 

            } // end for (;;)            

            GoOut: ;
            // encerra tudo
            vm_state_thread = false;
            this.ClosePort();
        }

        public bool ThRead_State()
        {
            // repassa se o serviço ainda está executando
            return vm_state_thread;
        }
        
        #endregion

        #region Cria a mensagem de dados do paciente para o aparelho

        public void queryresp()
        {
            // alerta
            InvokeIfRequired(com_bar, X => { com_bar.Text = "Modo Enquiry iniciado..."; });

            cReturnDataTable tmprec;
            DataRow dtrow;
            labmax_Item itemdetalhe;
            string codigobarras = "";
            string queryresp = "";
            int seqpac = 0;
            string paciente = "";
            int ord_seq_tx = 0;
            //string m_exames = "";
            string m_exa_cod = "";
            decimal m_qtd_exa = 0;
            //decimal m_qtd_amo = 0;
            //decimal m_qtd_reg = 0;
            //decimal m_tot_reg = 0;
            vm_frame = 0;
            string rack = "";
            int seq = 0;

            // solicita autorização para iniciar a transmissão
            int resultado = this.SendCommandToPort(enq);
            if (resultado != 0)
            {
                AddText("(Q) ENQ SendCommandToPort Error '" + resultado.ToString() + "'");
                return;
            }

            // Mensagem Header
            queryresp = "H";                                                        // Record Type ID
            queryresp = queryresp + fs + rd + cd + ed;                              // Delimiter Definition
            queryresp = queryresp + fs;                                             // Message Control ID
            queryresp = queryresp + fs;                                             // Access Password
            queryresp = queryresp + fs + hdr_rec_1;                                 // Sender Name or ID
            queryresp = queryresp + cd + hdr_rec_2;                                 //
            queryresp = queryresp + fs;                                             // Sender Street Address
            queryresp = queryresp + fs;                                             // Reserved Field
            queryresp = queryresp + fs;                                             // Sender Telephone Number
            queryresp = queryresp + fs;                                             // Characteristics of Sender
            queryresp = queryresp + fs + hdr_snd_1;                                 // Receiver ID
            queryresp = queryresp + cd + hdr_snd_2;                                 // Receiver ID
            queryresp = queryresp + fs;                                             // Comment or Special 
            queryresp = queryresp + fs + "P";                                       // Processing ID
            queryresp = queryresp + fs + "1";                                       // Version No.            
            queryresp = queryresp + fs + DateTime.Now.ToString("yyyyMMddHHmmss");   // data/hora no formato (year/month/day/hour/min/sec)              
            queryresp = queryresp + _cr;

            resultado = this.SendCommandToPort(queryresp);
            if (resultado != 0)
            {
                AddText("(Q) [H] SendCommandToPort Error '" + resultado.ToString() + "'");
                return;
            }

            // processa a lista de pedidos
            for (int x = 0; x < ExamesSolicitados.Count(); x++)
            {
                // obtem o item
                itemdetalhe = ExamesSolicitados.ListItem(x);
                rack = itemdetalhe.rack;

                if (codigobarras != itemdetalhe.amostra_barra)
                {
                    if (m_qtd_exa > 0)
                    {
                        ord_seq_tx++;
                        queryresp = "O";                                    // Record Type ID
                        queryresp = queryresp + fs + "1";                   // Sequence Number
                        queryresp = queryresp + fs + codigobarras;          // Specimen ID

                        if (rack != "0")
                            queryresp = queryresp + fs + cd + "1" + cd + seq.ToString(); // Instrument Specimen ID - Sequence
                        else
                            queryresp = queryresp + fs;                     // Instrument Specimen ID - Sequence

                        queryresp = queryresp + fs + m_exa_cod;             // Universal Test ID
                        queryresp = queryresp + fs + "R";                   // Priority
                        queryresp = queryresp + fs;                         // Requested/Ordered Date and 
                        queryresp = queryresp + fs;                         // Specimen Collection Date 
                        queryresp = queryresp + fs;                         // Collection End Time
                        queryresp = queryresp + fs;                         // Collection Volume
                        queryresp = queryresp + fs;                         // Collector ID
                        queryresp = queryresp + fs;                         // Action Code
                        queryresp = queryresp + fs;                         // Danger Code
                        queryresp = queryresp + fs;                         // Relevant Clinical 
                        queryresp = queryresp + fs;                         // Date/Time Specimen Received
                        queryresp = queryresp + fs + "Serum";               // Specimen Descriptor
                        queryresp = queryresp + fs;                         // Ordering Physician
                        queryresp = queryresp + fs;                         // Physician´s Telephone 
                        queryresp = queryresp + fs;                         // User Field No. 1
                        queryresp = queryresp + fs;                         // Users Field No. 2
                        queryresp = queryresp + fs;                         // Laboratory Field No. 1
                        queryresp = queryresp + fs;                         // Laboratory Field No. 2
                        queryresp = queryresp + fs;                         // Date/Time Results Reported 
                        queryresp = queryresp + fs;                         // Instrument Charge to 
                        queryresp = queryresp + fs;                         // Instrument Section ID 
                        queryresp = queryresp + fs + "O";                   // Report Types
                        queryresp = queryresp + _cr;

                        // envia Order
                        resultado = this.SendCommandToPort(queryresp);
                        if (resultado != 0)
                        {
                            AddText("(Q) [O] SendCommandToPort Error '" + resultado.ToString() + "'");
                            return;
                        }

                    } // (m_qtd_exa > 0)

                    codigobarras = itemdetalhe.amostra_barra;
                    seqpac++;
                    m_exa_cod = "";
                    m_qtd_exa = 0;
                    seq = itemdetalhe.seq;

                    // pesquisa o nome do paciente
                    SQL = "SELECT ";
                    SQL += "PACIENTE.PACIENTE, ";
                    SQL += "PACIENTE.NASCIMENTO, ";
                    SQL += "PACIENTE.SEXO ";
                    SQL += "FROM AMOSTRAS ";
                    SQL += "INNER JOIN PACIENTE ON ";
                    SQL += "PACIENTE.CODIGO=AMOSTRAS.CODIGO_PACIENTE ";
                    SQL += "WHERE AMOSTRAS.CODIGO_AMOSTRA='" + itemdetalhe.codigo_amostra + "'";
                    tmprec = mdb.Abrir_DataTable(SQL);
                    if (!tmprec.Status)
                    {
                        AddText("(Q) [P] Não conseguiu achar o paciente.");
                        return;
                    }

                    paciente = "NAO ENCONTRADO";

                    if (tmprec.RecordCount > 0)
                    {
                        dtrow = tmprec.dt.Rows[0];
                        paciente = dtrow["PACIENTE"].ToString();
                        if (paciente.Length > 20)
                            paciente = paciente.Substring(0, 19);
                    }
                    else
                    {
                        AddText("(Q) [P] Não conseguiu achar o paciente.");
                        return;
                    }

                    InvokeIfRequired(com_bar, X => { com_bar.Text = "Empacotando exames de '" + paciente + "'..."; });

                    // Mensagem Paciente
                    queryresp = "P";                                              // Record Type
                    queryresp = queryresp + fs + seqpac.ToString();               // Sequence Number    
                    queryresp = queryresp + fs + codigobarras;                    // Practice Assigned 
                    queryresp = queryresp + fs;                                   // 
                    queryresp = queryresp + fs;                                   // 
                    queryresp = queryresp + fs + Global.RemoverAcentos(paciente); // Patient Name
                    //queryresp = queryresp + fs;                                 // 
                    //queryresp = queryresp + fs;                                 // Date of birth
                    //queryresp = queryresp + fs;                                 // sex
                    //queryresp = queryresp + fs;                                 // 
                    //queryresp = queryresp + fs;                                 // 
                    //queryresp = queryresp + fs;                                 // 
                    //queryresp = queryresp + fs;                                 // 
                    //queryresp = queryresp + fs;                                 // doctor in charge
                    queryresp = queryresp + _cr;                                  //

                    // envia dados do paciente
                    resultado = this.SendCommandToPort(queryresp);
                    if (resultado != 0)
                    {
                        AddText("(Q) [P] SendCommandToPort Error '" + resultado.ToString() + "'");
                        return;
                    }

                    // adiciona um novo paciente e registra o codigo de barras
                    ord_seq_tx = 0;
                }

                // ORDER MESSAGE
                SQL = "SELECT CODIGO_ENVIO ";
                SQL += "FROM EQUIPAMENTOS_EXM_INTERFACE ";
                SQL += "WHERE CODIGO_EQUIPA=" + setup.ID.ToString() + " ";
                SQL += "AND CODIGO_EXAME=" + itemdetalhe.codigo_int + " ";
                SQL += "AND ACAO='Solicita Exame' ";
                SQL += "ORDER BY ORDEM";
                tmprec = mdb.Abrir_DataTable(SQL);

                foreach (DataRow dt in tmprec.dt.Rows)
                {
                    // cria os exames
                    m_exa_cod = m_exa_cod + (m_qtd_exa > 0 ? rd : "") 
                                          + cd + cd + cd
                                          + dt["CODIGO_ENVIO"].ToString().Trim() 
                                          + cd + cd 
                                          + "0";
                    m_qtd_exa++;
                }              

            } // proximo paciente

            if (m_qtd_exa > 0)
            {
                ord_seq_tx++;
                queryresp = "O";                                    // Record Type ID
                queryresp = queryresp + fs + ord_seq_tx.ToString(); // Sequence Number
                queryresp = queryresp + fs + codigobarras;          // Specimen ID

                if (rack != "0")
                    queryresp = queryresp + fs + cd + "1" + cd + seq.ToString(); // Instrument Specimen ID - Sequence
                else
                    queryresp = queryresp + fs;                     // Instrument Specimen ID - Sequence

                queryresp = queryresp + fs + m_exa_cod;             // Universal Test ID
                queryresp = queryresp + fs + "R";                   // Priority
                queryresp = queryresp + fs;                         // Requested/Ordered Date and 
                queryresp = queryresp + fs;                         // Specimen Collection Date 
                queryresp = queryresp + fs;                         // Collection End Time
                queryresp = queryresp + fs;                         // Collection Volume
                queryresp = queryresp + fs;                         // Collector ID
                queryresp = queryresp + fs;                         // Action Code
                queryresp = queryresp + fs;                         // Danger Code
                queryresp = queryresp + fs;                         // Relevant Clinical 
                queryresp = queryresp + fs;                         // Date/Time Specimen Received
                queryresp = queryresp + fs + "Serum";               // Specimen Descriptor
                queryresp = queryresp + fs;                         // Ordering Physician
                queryresp = queryresp + fs;                         // Physician´s Telephone 
                queryresp = queryresp + fs;                         // User Field No. 1
                queryresp = queryresp + fs;                         // Users Field No. 2
                queryresp = queryresp + fs;                         // Laboratory Field No. 1
                queryresp = queryresp + fs;                         // Laboratory Field No. 2
                queryresp = queryresp + fs;                         // Date/Time Results Reported 
                queryresp = queryresp + fs;                         // Instrument Charge to 
                queryresp = queryresp + fs;                         // Instrument Section ID 
                queryresp = queryresp + fs + "O";                   // Report Types
                queryresp = queryresp + _cr;

                if (queryresp.Length >= 240)
                {
                    int vm_f = 0;
                    string vm_f_querybuffer = "";

                    for (int f = 0; f < queryresp.Length; f++)
                    {
                        if (vm_f == 239)
                        {
                            // envia Order
                            if (f == queryresp.Length)
                            {
                                // ETX
                                resultado = this.SendCommandToPort(vm_f_querybuffer);
                                if (resultado != 0)
                                {
                                    AddText("(Q) [O] SendCommandToPort Error '" + resultado.ToString() + "'");
                                    return;
                                }
                            }
                            else
                            {
                                // ETB
                                resultado = this.SendCommandToPort(vm_f_querybuffer, true);
                                if (resultado != 0)
                                {
                                    AddText("(Q) [O] SendCommandToPort Error '" + resultado.ToString() + "'");
                                    return;
                                }
                            }

                            vm_f = 0;
                            vm_f_querybuffer = "";
                        }

                        vm_f_querybuffer += queryresp.Substring(f, 1);
                        vm_f++;
                    }

                    if (vm_f > 0)
                    {
                        // ETX
                        resultado = this.SendCommandToPort(vm_f_querybuffer);
                        if (resultado != 0)
                        {
                            AddText("(Q) [O] SendCommandToPort Error '" + resultado.ToString() + "'");
                            return;
                        }

                        vm_f = 0;
                        vm_f_querybuffer = "";
                    }
                }
                else
                {
                    // envia Order
                    resultado = this.SendCommandToPort(queryresp);
                    if (resultado != 0)
                    {
                        AddText("(Q) [O] SendCommandToPort Error '" + resultado.ToString() + "'");
                        return;
                    }
                }

            } // (m_qtd_exa > 0)

            // terminator escape
            queryresp = "L";
            queryresp = queryresp + fs + "1";
            queryresp = queryresp + fs + "N";
            queryresp = queryresp + _cr;

            // envia 
            resultado = this.SendCommandToPort(queryresp);
            if (resultado != 0)
            {
                AddText("(Q) [L] SendCommandToPort Error '" + resultado.ToString() + "'");
                return;
            }

            // envia EOT para finalizar a transmissão
            resultado = this.SendCommandToPort(eot);
            if (resultado != 0)
            {
                AddText("(Q) ENQ SendCommandToPort Error '" + resultado.ToString() + "'");
                return;
            }

        } // fim de função

        #endregion 

        #region Sistema de Cálculos

        public decimal PreparaFormula(string pre_calculo, string codigo_amostra, string codigo_equipa, string codigo_exame)
        {
            // pega o calculo ainda com máscaras e substitui pelos valores verdadeiros

            string retorno = pre_calculo;

            SQL = "SELECT ";
            SQL += "EQUIPAMENTOS_EXM_INTERFACE.VAR_CALCULO, ";
            SQL += "INTERFACE_CAMPOS.R_RESULT ";
            SQL += "FROM INTERFACE_CAMPOS ";
            SQL += "INNER JOIN EQUIPAMENTOS_EXM_INTERFACE ON ";
            SQL += "EQUIPAMENTOS_EXM_INTERFACE.ID=INTERFACE_CAMPOS.ID_CAMPO ";
            SQL += "WHERE INTERFACE_CAMPOS.CODIGO_AMOSTRA='" + codigo_amostra + "' ";
            SQL += "AND INTERFACE_CAMPOS.CODIGO_EQUIPA=" + codigo_equipa + " ";
            SQL += "AND INTERFACE_CAMPOS.CODIGO_EXAME=" + codigo_exame + " ";
            SQL += "ORDER BY  ";
            SQL += "INTERFACE_CAMPOS.R_ORDER, ";
            SQL += "INTERFACE_CAMPOS.R_SEQ ";
            cReturnDataTable tmprec = mdb.Abrir_DataTable(SQL);
            if (!tmprec.Status)
                return 0;

            // troca as variaveis pelos valores reais
            foreach (DataRow dt in tmprec.dt.Rows)
            {
                if (dt["VAR_CALCULO"].ToString().Trim().Length > 0)
                {
                    retorno = retorno.Replace(dt["VAR_CALCULO"].ToString(), dt["R_RESULT"].ToString());
                }
            }

            // calcula realmente
            return (retorno.Length > 0) ? this.Calculate(retorno) : 0;
        }

        public string FormatEx(string input, int qtdecimal)
        {
            // formata resultado de acordo com qtd de casas decimais

            Misc misc = new Misc();
            string format = "0"; // #.###0

            for (int x = 1; x < qtdecimal; x++)
            {
                format = "#" + format;
            }

            format = "0." + format;

            return misc.FormatMoney(input, format);
        }

        public decimal Calculate(string formula)
        {
            decimal result = 0;

            try
            {
                // metodo B
                // base projeto: http://ncalc.codeplex.com/
                // funções: http://ncalc.codeplex.com/wikipage?title=functions&referringTitle=Home
                //

                Expression runFormulaB = new Expression(formula);
                result = decimal.Parse(runFormulaB.Evaluate().ToString());
                runFormulaB = null;
            }
            catch { }

            return result;
        }

        public void ProccessCalculate(string codigo_barra)
        {
            // função lê o codigo de barras da amostra e
            // processa todos os calculos nos resultados
            // recebidos pelo equipamento.

            // separa a string
            string setor = "";
            string codigoposto = "";
            string amostrareal = "";
            cReturnDataTable tmprec;
            cInsertInto inc = new cInsertInto();

            try
            {
                setor = codigo_barra.Substring(0, 3);
                codigoposto = codigo_barra.Substring(3, 3);
                amostrareal = codigo_barra.Substring(6, codigo_barra.Length - 6);
                amostrareal = int.Parse(amostrareal).ToString();
            }
            catch
            {
                setor = "";
                codigoposto = "";
                amostrareal = "";
                AddText("Erro: amostra não encontrada!");
                return;
            }

            if (setor.Trim() == "" || codigoposto.Trim() == "" || amostrareal.Trim() == "")
            {
                AddText("Erro: amostra não encontrada!");
                return;
            }

            // levanta os exames que deverão ser feitos
            SQL = "SELECT ";
            SQL += "AMOSTRAS.CODIGO_AMOSTRA, ";
            SQL += "AMOSTRAS_EXAMES.CODIGO_EXAME, ";
            SQL += "AMOSTRAS_EXAMES.SEQUENCIA, ";
            SQL += "EXAMES.CODIGO, ";
            SQL += "PACIENTE.NASCIMENTO, ";
            SQL += "PACIENTE.SEXO ";
            SQL += "FROM AMOSTRAS ";
            SQL += "INNER JOIN AMOSTRAS_EXAMES ON ";
            SQL += "AMOSTRAS.CODIGO_AMOSTRA=AMOSTRAS_EXAMES.CODIGO_AMOSTRA ";
            SQL += "INNER JOIN EXAMES ON ";
            SQL += "EXAMES.CODIGO_INT=AMOSTRAS_EXAMES.CODIGO_EXAME ";
            SQL += "INNER JOIN PACIENTE ON ";
            SQL += "PACIENTE.CODIGO=AMOSTRAS.CODIGO_PACIENTE ";
            SQL += "WHERE AMOSTRAS.CODIGO_POSTOCOLETA=" + int.Parse(codigoposto).ToString() + " ";
            SQL += "AND AMOSTRAS.AMOSTRA='" + amostrareal + "' ";
            SQL += "AND EXAMES.GER_SETCOD=" + setor + " ";
            tmprec = mdb.Abrir_DataTable(SQL);
            if (!tmprec.Status)
                return;

            cReturnDataTable campos;
            string resultado = "";
            decimal dec_resultado = 0;

            foreach (DataRow dt in tmprec.dt.Rows)
            {
                // levanta os campos dos exames
                SQL = "SELECT ";
                SQL += "EQUIPAMENTOS_EXM_INTERFACE.DECIMAL,";
                SQL += "EQUIPAMENTOS_EXM_INTERFACE.ACAO, ";
                SQL += "EQUIPAMENTOS_EXM_INTERFACE.CALCULO, ";
                SQL += "INTERFACE_CAMPOS.R_RESULT, ";
                SQL += "INTERFACE_CAMPOS.ID_CAMPO, ";
                SQL += "INTERFACE_CAMPOS.R_ORDER, ";
                SQL += "INTERFACE_CAMPOS.R_SEQ ";
                SQL += "FROM INTERFACE_CAMPOS ";
                SQL += "INNER JOIN EQUIPAMENTOS_EXM_INTERFACE ON ";
                SQL += "EQUIPAMENTOS_EXM_INTERFACE.ID=INTERFACE_CAMPOS.ID_CAMPO ";
                SQL += "WHERE INTERFACE_CAMPOS.CODIGO_AMOSTRA='" + dt["CODIGO_AMOSTRA"].ToString() + "' ";
                SQL += "AND INTERFACE_CAMPOS.CODIGO_EQUIPA=" + setup.ID.ToString() + " ";
                SQL += "AND INTERFACE_CAMPOS.CODIGO_EXAME=" + dt["CODIGO_EXAME"].ToString() + " ";
                SQL += "ORDER BY ";
                SQL += "INTERFACE_CAMPOS.CODIGO_EXAME, ";
                SQL += "INTERFACE_CAMPOS.R_ORDER, ";
                SQL += "INTERFACE_CAMPOS.R_SEQ ";
                campos = mdb.Abrir_DataTable(SQL);
                if (!campos.Status)
                    return;

                foreach (DataRow dtcalc in campos.dt.Rows)
                {
                    resultado = dtcalc["R_RESULT"].ToString();

                    if (dtcalc["ACAO"].ToString() == "Calcula")
                    {
                        if (dtcalc["CALCULO"].ToString().Trim().Length > 0)
                        {
                            // gera calculo
                            dec_resultado = PreparaFormula(dtcalc["CALCULO"].ToString(),
                                                           dt["CODIGO_AMOSTRA"].ToString(),
                                                           setup.ID.ToString(),
                                                           dt["CODIGO_EXAME"].ToString()
                                                           );

                            resultado = dec_resultado.ToString();

                            // formata resultado
                            if (int.Parse(dtcalc["DECIMAL"].ToString()) > 0)
                                resultado = FormatEx(dec_resultado.ToString(), int.Parse(dtcalc["DECIMAL"].ToString()));

                            // atualiza resultado base de dados
                            SQL = "WHERE INTERFACE_CAMPOS.CODIGO_AMOSTRA='" + dt["CODIGO_AMOSTRA"].ToString() + "' ";
                            SQL += "AND INTERFACE_CAMPOS.CODIGO_EQUIPA=" + setup.ID.ToString() + " ";
                            SQL += "AND INTERFACE_CAMPOS.CODIGO_EXAME=" + dt["CODIGO_EXAME"].ToString() + " ";
                            SQL += "AND R_ORDER=" + dtcalc["R_ORDER"].ToString() + " ";
                            SQL += "AND R_SEQ=" + dtcalc["R_SEQ"].ToString() + " ";

                            resultado = resultado.Replace(",", ".");

                            inc.Clear();
                            inc.Add("R_RESULT", resultado, true);
                            inc.Add("R_REF", (dtcalc["CALCULO"].ToString().Length > 100) ? dtcalc["CALCULO"].ToString().Substring(0, 100) : dtcalc["CALCULO"].ToString(), true);
                            SQL = inc.Update("INTERFACE_CAMPOS", SQL);
                            if (!mdb.SQLExecute(SQL))
                                return;
                        }
                    } // end -> validação para campo de Calculo
                } // end -> dtcalc
            } // end -> dt

        } // end function()

        #endregion

    }
}