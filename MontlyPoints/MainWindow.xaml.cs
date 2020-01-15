using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using Newtonsoft.Json;

namespace MontlyPoints
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
	{
        #region Declarations

        public int NivelPontualidade = 5;
		public int HoraEntrada { get; set; }
		public int MinutosEntrada { get; set; }
		public int HoraSaida { get; set; }
		public int MinutosSaida { get; set; }
		public string DiaInicial { get; set; }
		public int TotalLinhas { get; set; }
        public DateTime DataInicial { get; set; }
        public DateTime DataFinal { get; set; }

        #endregion

        #region Events

        public MainWindow()
		{
			InitializeComponent();
			CarregarValoresPadrao();
		}

		private void Button_Click(object sender, RoutedEventArgs e)
		{
			GerarPlanilha();
		}

		private void TxtHora_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
		{

		}

		#endregion

		#region Methods

        /// <summary>
        /// Carrega os valores-padrão do programa
        /// </summary>
		private void CarregarValoresPadrao()
		{
            DateTime dataHoje = DateTime.Today;
            string diaFinal = string.Empty;

            //Recupera os valores do app.config
            DiaInicial = System.Configuration.ConfigurationManager.AppSettings["diaInicial"].ToString();
            diaFinal = System.Configuration.ConfigurationManager.AppSettings["diaFinal"].ToString();

            //Calcula a dataInicial, pegando a data do app config e aplicando no mês passado
            DataInicial = new DateTime(dataHoje.AddMonths(-1).Year,
                        dataHoje.AddMonths(-1).Month, int.Parse(DiaInicial));

            //Calcula a dataFinal, pegando a data do app config e aplicando no mês atual
            DataFinal = new DateTime(dataHoje.Year,
                        dataHoje.Month, int.Parse(diaFinal));

            //Calcula o total de linhas
            TotalLinhas = Convert.ToInt32((DataFinal - DataInicial).TotalDays) + 16;

            //Seta os valores-padrão dos controles
            txtHora.Text = "09";
			txtMinutos.Text = "00";
		}

		/// <summary>
		/// Gera e abre a planilha
		/// </summary>
		private void GerarPlanilha()
		{
			using (new WaitCursor())
			{
				List<MDLFeriado> feriados = new List<MDLFeriado>();
				FileInfo fileInfo;
				ExcelPackage excelPackage;
				ExcelWorksheet planilha;
				Random random = new Random();
				DateTime dataContador = DateTime.Today;
				string pathNovoDocumento = string.Empty;
				int diasNoMes = int.MinValue;
				int diferenca = int.MinValue;

				//Se campos forem válidos
				if (ValidarCampos())
				{
					//Obtem os horarios de entradaSaida
					CalcularHoraEntradaSaida();

					//Busca os feriados na API
					feriados = RetornarFeriados();

					//Obtêm o novoDocumento
					fileInfo = PrepararArquivo(out pathNovoDocumento);
					excelPackage = new ExcelPackage(fileInfo);

					//Obtêm a planilha
					planilha = excelPackage.Workbook.Worksheets["planilha"];

					//Preenche o cabeçalho da planilha
					PreencherCabecalho(planilha);

					//Calcula quantos dias o mês passado teve 28/29/30/31
					diasNoMes = DateTime.DaysInMonth(DateTime.Today.AddMonths(-1).Year,
                                    DateTime.Today.AddMonths(-1).Month);

					//Se for menos de 31 dias tira as linhas que sobraram
					if (diasNoMes < 31)
					{
						diferenca = 31 - diasNoMes;
						planilha.DeleteRow(diasNoMes - 5, diferenca, true);
						TotalLinhas -= diferenca;
					}

                    //Atribui a variável que será usada como contador
                    dataContador = DataInicial;

                    //Começa a preencher as linhas do grid, o grid começa na linha 15
                    for (int counter = 15; counter < TotalLinhas; counter++)
					{
                        PreencherDias(planilha, dataContador, counter);

                        //Se for fim de semana ou feriado pinta a linha
                        if
						(
                            dataContador.DayOfWeek == DayOfWeek.Sunday ||
                            dataContador.DayOfWeek == DayOfWeek.Saturday ||
							feriados.Any(feriado => 
                            feriado.DataFeriado.Date == dataContador.Date
                            )
						)
						{
							PintarLinha(planilha, counter);
						}
						//Senão preenche os horários
						else
						{
							PreencherEntradaSaida(planilha, counter, random);
							PreencherIntervalo2(planilha, counter, random);
						}

                        dataContador = dataContador.AddDays(1);
					}

                    planilha = LimparLinhasSobrando(planilha);

					excelPackage.Save();

					Process.Start(pathNovoDocumento);
				}
			}
		}

        /// <summary>
        /// Deleta as linhas que estão sobrando na planilha
        /// </summary>
        /// <param name="planilha"></param>
        private ExcelWorksheet LimparLinhasSobrando(ExcelWorksheet planilha)
        {
            var inicioPlanilha = planilha.Dimension.Start;
            var fimPlanilha = planilha.Dimension.End;

            //Itera pelas linhas da planilha
            for (int row = inicioPlanilha.Row; row <= fimPlanilha.Row; row++)
            {
                //Se dia for 0 deleta a linha
                if(planilha.Cells[row, 1].Text == "0")
                {
                    planilha.DeleteRow(row);
                    row--;
                }
            }

            return planilha;
        }

        /// <summary>
        /// Valida os inputs
        /// </summary>
        private bool ValidarCampos()
		{
			bool valido = true;

			if (string.IsNullOrWhiteSpace(txtMinutos.Text) || string.IsNullOrWhiteSpace(txtHora.Text) || string.IsNullOrWhiteSpace(txtNome.Text)
				|| string.IsNullOrWhiteSpace(txtfuncao.Text) || string.IsNullOrWhiteSpace(txctCpf.Text) || string.IsNullOrWhiteSpace(txtMatricula.Text))
			{
				MessageBox.Show("Informe todos os campos para prosseguir");
				valido = false;
			}

			return valido;
		}

		/// <summary>
		/// Retorna uma lista com os feriados do ano
		/// </summary>
		/// <returns></returns>
		private List<MDLFeriado> RetornarFeriados()
		{
			// Eric.Wu: Implementada base de feriados do fasys 

			List<MDLFeriado> listaFeriados= new List<MDLFeriado>();
			string pastaParente = Directory.GetParent(Assembly.GetExecutingAssembly().Location.ToString()).ToString();

			listaFeriados = JsonConvert.DeserializeObject<List<MDLFeriado>>(File.ReadAllText(Path.Combine(pastaParente, "Feriados.json")));

			listaFeriados = listaFeriados.Where(feriado => feriado.DataFeriado.Year == DateTime.Now.Year || feriado.DataFeriado.Year == (DateTime.Now.Year - 1)).ToList();

			return listaFeriados;
		}

		///// <summary>
		///// Retorna uma lista com os feriados do ano
		///// </summary>
		///// <returns></returns>
		//private List<Feriado> RetornarFeriados()
		//{
		//	List<Feriado> feriados = new List<Feriado>();
		//	List<Feriado> feriadosFinal = new List<Feriado>();
		//	string urlAPI = "https://api.calendario.com.br/?json=true&ano=2019&ibge=3548708&token=dmluaWNpdXMubWFyaWFub0BmYWNpbGFzc2lzdC5jb20uYnImaGFzaD03MDk4MDYzNg";

		//	try
		//	{
		//		//Chama a integração com a API para obter feriados
		//		var client = new RestClient(urlAPI);
		//		var request = new RestRequest("", Method.GET) { RequestFormat = RestSharp.DataFormat.Json };
		//		IRestResponse response = client.Execute(request);
		//		var content = response.Content;

		//		//Deserializa os feriados
		//		feriados = new System.Web.Script.Serialization.JavaScriptSerializer()
		//			.Deserialize<List<Feriado>>(content.ToString());

		//		//Formata as datas
		//		foreach (Feriado feriado in feriados)
		//		{
		//			feriado.dateFormated = DateTime.ParseExact(feriado.date, "dd/MM/yyyy", CultureInfo.InvariantCulture);
		//			feriadosFinal.Add(feriado);
		//		}
		//	}
		//	catch(Exception ex)
		//	{}

		//	return feriadosFinal;
		//}
		
		/// <summary>
		/// Cria o novo arquivo a ser preenchido
		/// </summary>
		/// <param name="pathNovoDocumento"></param>
		/// <returns></returns>
        private FileInfo PrepararArquivo(out string pathNovoDocumento)
        {
			string pastaParente = string.Empty;
            string pathTemplate = string.Empty;
			DirectoryInfo directoryInfo;

			//Obtem a pasta em que o programa está sendo executado
			pastaParente = Directory.GetParent(Assembly.GetExecutingAssembly().Location.ToString()).ToString();

            //Cria pasta se não existir
            directoryInfo = Directory.CreateDirectory(string.Format("{0}/documentos/", pastaParente));

            //Obtem os paths
            pathTemplate = string.Format("{0}\\template.xlsx", pastaParente);
            pathNovoDocumento = string.Format("{0}/documentos/{1}.xlsx", pastaParente, DateTime.Now.ToString("ddMMyyyyHHmmss"));

			//Gera o novoDocumento
			File.Copy(pathTemplate, pathNovoDocumento, true);

            return new FileInfo(pathNovoDocumento);
        }

		/// <summary>
		/// Preenche o cabeçalho da planilha
		/// </summary>
		/// <param name="myWorksheet"></param>
        private void PreencherCabecalho(ExcelWorksheet myWorksheet)
        {
			DateTime dataSelecionada = new DateTime();
			string mesPassado = string.Empty;
			string mesAtual = string.Empty;
			string horarioTrabalho = string.Empty;
			string minutesStr = string.Empty;

			//Calcula os minutos de entrada
			minutesStr = MinutosEntrada.ToString().Length < 2 ? "0" + MinutosEntrada.ToString() : MinutosEntrada.ToString();

			//Monta o horário de trabalho
			horarioTrabalho = string.Format("{0}:{1} às {2}:{1}", HoraEntrada, minutesStr, HoraSaida);

			//Recupera o texto dos meses
			dataSelecionada = DateTime.Today;
			mesPassado = ConverterMes(dataSelecionada.AddMonths(-1).Month);
			mesAtual = ConverterMes(dataSelecionada.Month);

			myWorksheet.Cells[8, 1].Value = txtNome.Text;
            myWorksheet.Cells[8, 10].Value = txctCpf.Text;
            myWorksheet.Cells[6, 10].Value = txtMatricula.Text;
            myWorksheet.Cells[10, 1].Value = txtfuncao.Text;      			
            myWorksheet.Cells[10, 9].Value = horarioTrabalho; //Hora de Trabalho
			myWorksheet.Cells[4, 1].Value = string.Format("{0}/{1}", mesPassado.ToLower(), mesAtual.ToLower());
            myWorksheet.Cells[4, 3].Value = dataSelecionada.Year;
        }

		/// <summary>
		/// Converte o int mes em string
		/// </summary>
		/// <param name="mes"></param>
		/// <returns></returns>
		private string ConverterMes(int mes)
        {
            string mesConvertido = "";

            switch (mes)

            {
                default:
                case 1:
                    mesConvertido = "Janeiro";
                    break;

                case 2:
                    mesConvertido = "Fevereiro";
                    break;

                case 3:
                    mesConvertido = "Março";
                    break;

                case 4:
                    mesConvertido = "Abril";
                    break;

                case 5:
                    mesConvertido = "Maio";
                    break;

                case 6:
                    mesConvertido = "Junho";
                    break;

                case 7:
                    mesConvertido = "Julho";
                    break;

                case 8:
                    mesConvertido = "Agosto";
                    break;

                case 9:
                    mesConvertido = "Setembro";
                    break;

                case 10:
                    mesConvertido = "Outubro";
                    break;

                case 11:
                    mesConvertido = "Novembro";
                    break;

                case 12:
                    mesConvertido = "Dezembro";
                    break;
            }

            return mesConvertido;
        }

        #region PreencherLinhas    

        /// <summary>
        /// Preenche os pontos de entrada e saída
        /// </summary>
        /// <param name="myWorksheet"></param>
        /// <param name="counter"></param>
        /// <param name="random"></param>
        private void PreencherDias(ExcelWorksheet myWorksheet, DateTime data, int counter)
        {
            //Preenche a célula de dia
            myWorksheet.Cells[counter, 1].Value = data.Day;              
        }

        /// <summary>
        /// Preenche os horários de intervalo
        /// </summary>
        /// <param name="myWorksheet"></param>
        /// <param name="counter"></param>
        /// <param name="rn"></param>
        private void PreencherIntervalo2(ExcelWorksheet myWorksheet, int counter, Random rn)
        {
			string minutesStr = string.Empty;
			int minutes = int.MinValue;

			minutes = rn.Next(0, NivelPontualidade);
            minutesStr = minutes.ToString().Length < 2 ? "0" + minutes.ToString() : minutes.ToString();

            myWorksheet.Cells[counter, 5].Value = string.Format("12:{0}", minutesStr);
            myWorksheet.Cells[counter, 6].Value = string.Format("13:{0}", minutesStr);
        }

        /// <summary>
        /// Preenche os pontos de entrada e saída
        /// </summary>
        /// <param name="myWorksheet"></param>
        /// <param name="counter"></param>
        /// <param name="random"></param>
        private void PreencherEntradaSaida(ExcelWorksheet myWorksheet, int counter, Random random)
        {
            string minutosEmTexto = string.Empty;
            int minutos = int.MinValue;
			int variacaoMinutos = int.MinValue;

			//Calcula a máxima variação dos minutos
			variacaoMinutos = MinutosEntrada + NivelPontualidade > 59 ? 59 : MinutosEntrada + NivelPontualidade;

			//Calcula o minutos e coloca em uma string
			minutos = random.Next(MinutosEntrada, variacaoMinutos);            
            minutosEmTexto = minutos.ToString().Length < 2 ? "0" + minutos.ToString() : minutos.ToString();

            //Preenche as células de entrada e saída
            myWorksheet.Cells[counter, 2].Value = string.Format("{0}:{1}", HoraEntrada, minutosEmTexto);
            myWorksheet.Cells[counter, 9].Value = string.Format("{0}:{1}", HoraSaida, minutosEmTexto);
        }

        /// <summary>
        /// Calcula as horas de entrada e saída
        /// </summary>
        private void CalcularHoraEntradaSaida()
        {
            int horaEntrada = int.MinValue;
            int minutosEntrada = int.MinValue;

            //Pega a hora de entrada especificada
            int.TryParse(txtHora.Text, out horaEntrada);
            int.TryParse(txtMinutos.Text, out minutosEntrada);

            HoraEntrada = horaEntrada;
            MinutosEntrada = minutosEntrada;

            //Calcula a hora de saída
            HoraSaida = horaEntrada + 9;
        }

        /// <summary>
        /// Preenche as colunas de uma determinada linha
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="counter"></param>
        private void PintarLinha(ExcelWorksheet excelWorksheet, int counter)
        {
            for (int coluna = 1; coluna < 14; coluna++)
            {
                excelWorksheet.Cells[counter, coluna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                excelWorksheet.Cells[counter, coluna].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
            }
        }

        #endregion
        #endregion
    }
}