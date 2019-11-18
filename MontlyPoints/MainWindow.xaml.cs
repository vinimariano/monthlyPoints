using OfficeOpenXml;
using OfficeOpenXml.Style;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Input;
using Newtonsoft.Json;

namespace MontlyPoints
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		#region Declarations

		public int NivelPontualidade { get; set; }
		public int HoraEntrada { get; set; }
		public int MinutosEntrada { get; set; }
		public int HoraSaida { get; set; }
		public int MinutosSaida { get; set; }

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

		private void CarregarValoresPadrao()
		{
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
				DateTime dataCorrente = DateTime.MinValue;
				Random random = new Random();
				DateTime dataHoje = DateTime.Today;
				string pathNovoDocumento = string.Empty;
				int totalLinhas = int.MinValue;
				int diasNoMes = int.MinValue;
				int diferenca = int.MinValue;

				//Se campos forem válidos
				if (ValidarCampos())
				{
					//Multiplica o valor do slider por 10
					NivelPontualidade = Convert.ToInt32(SliderPontualidade.Value) * 10;

					//Obtem os horarios de entradaSaida
					CalcularHoraEntradaSaida();

					//Busca os feriados na API
					feriados = RetornarFeriados();

					//Obtêm o total das linhas
					totalLinhas = 46;

					//Obtêm o novoDocumento
					fileInfo = PrepararArquivo(out pathNovoDocumento);
					excelPackage = new ExcelPackage(fileInfo);

					//Obtêm a planilha
					planilha = excelPackage.Workbook.Worksheets["planilha"];

					//Calcula a data inicial
					dataCorrente = new DateTime(dataHoje.AddMonths(-1).Year,
						dataHoje.AddMonths(-1).Month, 21);

					//Preenche o cabeçalho da planilha
					PreencherCabecalho(planilha);

					//Calcula os dias no mês
					diasNoMes = DateTime.DaysInMonth(dataHoje.AddMonths(-1).Year,
									dataHoje.AddMonths(-1).Month);

					//Se for menos de 31 dias tira as linhas que sobraram
					if (diasNoMes < 31)
					{
						diferenca = 31 - diasNoMes;
						planilha.DeleteRow(diasNoMes - 5, diferenca, true);
						totalLinhas = 46 - diferenca;
					}

					//Começa a preencher as linhas do grid, o grid começa na linha 15
					for (int counter = 15; counter < totalLinhas; counter++)
					{
						//Se for fim de semana ou feriado pinta a linha
						if
						(
							dataCorrente.DayOfWeek == DayOfWeek.Sunday ||
							dataCorrente.DayOfWeek == DayOfWeek.Saturday ||
							feriados.Any(feriado => feriado.DataFeriado.Date == dataCorrente.Date)
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

						dataCorrente = dataCorrente.AddDays(1);
					}

					excelPackage.Save();

					Process.Start(pathNovoDocumento);
				}
			}
		}

		/// <summary>
		/// Valida os inputs
		/// </summary>
		private bool ValidarCampos()
		{
			bool valido = true;

			if (string.IsNullOrWhiteSpace(txtMinutos.Text) || string.IsNullOrWhiteSpace(txtHora.Text) || string.IsNullOrWhiteSpace(txtNome.Text)
				|| string.IsNullOrWhiteSpace(txtfuncao.Text) || string.IsNullOrWhiteSpace(txctCpf.Text))
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

			listaFeriados = listaFeriados.Where(feriado => feriado.DataFeriado.Year == DateTime.Now.Year).ToList();

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
            pathTemplate = string.Format("{0}//template.xlsx", pastaParente);
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

			myWorksheet.Cells[8, 2].Value = txtNome.Text;
            myWorksheet.Cells[8, 10].Value = txctCpf.Text;
            myWorksheet.Cells[10, 1].Value = txtfuncao.Text;      			
            myWorksheet.Cells[10, 9].Value = horarioTrabalho; //Hora de Trabalho
			myWorksheet.Cells[4, 1].Value = string.Format("{0}/{1}", mesPassado.ToLower(), mesAtual.ToLower());
            myWorksheet.Cells[4, 3].Value = dataSelecionada.Year;
            myWorksheet.Cells[4, 3].Value = dataSelecionada.Year;
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
		/// Converte o int mes em string
		/// </summary>
		/// <param name="mes"></param>
		/// <returns></returns>
		private string ConverterMes(int mes)
        {
            string mesConvertido = "";

            switch (mes)

            {
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

                default:
                    mesConvertido = "Janeiro";
                    break;
            }

            return mesConvertido;
        }

		/// <summary>
		/// Preenche os horários de intervalo
		/// </summary>
		/// <param name="myWorksheet"></param>
		/// <param name="counter"></param>
		/// <param name="rn"></param>
        private void PreencherIntervalo2(ExcelWorksheet myWorksheet, int counter, Random rn)
        {
			string minutesstr = string.Empty;
			int minutes = int.MinValue;

			minutes = rn.Next(0, NivelPontualidade);
            minutesstr = minutes.ToString().Length < 2 ? "0" + minutes.ToString() : minutes.ToString();

            myWorksheet.Cells[counter, 5].Value = string.Format("12:{0}", minutesstr);
            myWorksheet.Cells[counter, 6].Value = string.Format("13:{0}", minutesstr);
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
	}
}