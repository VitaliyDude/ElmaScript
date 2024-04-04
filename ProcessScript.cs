using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EleWise.ELMA.API;
using EleWise.ELMA.Model.Common;
using EleWise.ELMA.Model.Entities;
using EleWise.ELMA.Model.Managers;
using EleWise.ELMA.Model.Types.Settings;
using EleWise.ELMA.Model.Entities.ProcessContext;
using Context = EleWise.ELMA.Model.Entities.ProcessContext.P_Pokazateli1S;

namespace EleWise.ELMA.Model.Scripts
{
	/// <summary>
	/// Модуль сценариев процесса "Показатели 1С"
	/// </summary>
	/// <example> 
	/// <![CDATA[
	/// >>>>>>>>>>>>>>>ВАЖНАЯ ИНФОРМАЦИЯ!!!<<<<<<<<<<<<<<<
	/// Данный редактор создан для работы с PublicAPI. 
	/// PublicAPI предназначен для разработки сценариев ELMA.
	/// Например, с помощью PublicAPI можно добавить комментарий к документу:
	/// //Загружаем документ
	/// var doc = PublicAPI.Docflow.Document.Load(56);
	/// //Добавляем комментарий
	/// PublicAPI.Docflow.Document.AddComment(doc, "тут ваш комментарий");
	/// 
	/// Более подробно про PublicAPI вы можете узнать тут: http://www.elma-bpm.ru/kb/article-642ApiRoot.html
	/// 
	/// Если же вам нужна более серьёзная разработка, выходящая за рамки PublicAPI, используйте
	/// сторонние редакторы кода, такие как SharpDevelop и VisualStudio.
	/// Информацию по запуску кода в стороннем редакторе вы можете найти тут:
	/// http://www.elma-bpm.ru/kb/article-837.html
	/// ]]>
	/// </example>
	public partial class P_Pokazateli1S_Scripts : EleWise.ELMA.Workflow.Scripts.ProcessScriptBase<Context>
	{
		#region Константы
		/// <summary>
		/// Строка подключения для 1с
		/// </summary>
		//const string ONE_C_CONNECTOR = "1C-СМК"; //Test
		const string ONE_C_CONNECTOR = "1C-RD";
		const long PROCESS_HEADER_ID = 2;

		// Prod
		/// <summary>
		/// Тег лаборатории
		/// </summary>
		const string LABORATORY_TAG = "НЛ ";

		#region Кадровый состав лаборатории
		/// <summary>
		/// Количество сотрудников лаборатории
		/// </summary>
		const string METRIC_COUNT_SOTRUDNIK = "Количество сотрудников лаборатории";

		/// <summary>
		/// Количество исследователей, направленных на работу (стажировку) в зарубежные организации (научные организации/университеты)
		/// </summary>
		const string METRIC_COUNT_TRAVELERS = "Количество исследователей, направленных на работу (стажировку) в зарубежные организации (научные организации/университеты)";

		/// <summary>
		/// Доля исследователей в составе лаборатории в возрасте до 39 лет
		/// </summary>
		const string METRIC_SHARE_SOTRUDNIK_BEFORE_39 = "Доля исследователей в составе лаборатории в возрасте до 39 лет";

		/// <summary>
		/// Доля сотрудников лаборатории с наличием степени кандидата наук
		/// </summary>
		const string METRIC_SHARE_SOTRUDNIK_CANDIDATE = "Доля сотрудников лаборатории с наличием степени кандидата наук";

		/// <summary>
		/// Доля сотрудников лаборатории с наличием степени доктора наук
		/// </summary>
		const string METRIC_SHARE_SOTRUDNIK_DOCTOR = "Доля сотрудников лаборатории с наличием степени доктора наук";

		/// <summary>
		/// Доля НПР кафедр, трудоустроенных в лабораторию, в общей численности лаборатории
		/// </summary>
		const string METRIC_SHARE_SOTRUDNIK_NPR_DEPARTMENT = "Доля НПР кафедр, трудоустроенных в лабораторию, в общей численности лаборатории";

		/// <summary>
		/// Перечень защищенных диссертаций сотрудниками лаборатории (кандидатские)
		/// </summary>
		const string METRIC_LIST_CANDIDATE_WORK = "Перечень защищенных диссертаций сотрудниками лаборатории (кандидатские)";

		/// <summary>
		/// Перечень защищенных диссертаций сотрудниками лаборатории (докторские)
		/// </summary>
		const string METRIC_LIST_DOCTORAL_WORK = "Перечень защищенных диссертаций сотрудниками лаборатории (докторские)";
		
		/// <summary>
		/// Перечень оборудования лаборатории
		/// </summary>
		const string METRIC_EQUIPMENT_LIST_EQUIPMENT = "Перечень оборудования лаборатории";
		
		/// <summary>
		/// Перечень используемых технологий
		/// </summary>
		const string METRIC_TECHNOLOGY_LIST_TECHNOLOGY = "Перечень используемых технологий";
		
		/// <summary>
		/// Количество обучающихся (студенты СПО, студенты ВО), прошедших (проходящих) практику в лаборатории
		/// </summary>
		const string METRIC_STUDY_LIST_STUDENTS = "Количество обучающихся (студенты СПО, студенты ВО), прошедших (проходящих) практику в лаборатории";
		
		/// <summary>
		/// Перечень образовательных модулей, которые ведут сотрудники лаборатории
		/// </summary>
		const string METRIC_STUDY_LIST_STUDY_MODULES = "Перечень образовательных модулей, которые ведут сотрудники лаборатории";
		
		/// <summary>
		/// Список патентов на изобретение, полезную модель, полученных в ходе работ, проводимых в лаборатории
		/// </summary>
		const string METRIC_SCIENCE_RID_PATTENT = "Список патентов на изобретение, полезную модель, полученных в ходе работ, проводимых в лаборатории";
		
		/// <summary>
		/// Список свидетельств о гос.регистрации программы для ЭВМ, баз данных, полученных в ходе работ, проводимых в лаборатории
		/// </summary>
		const string METRIC_SCIENCE_RID_SERTIFICATE = "Список свидетельств о гос.регистрации программы для ЭВМ, баз данных, полученных в ходе работ, проводимых в лаборатории";
		
		/// <summary>
		/// Доля дохода от лицензионных соглашений от объема внебюджетных средств лаборатории в год
		/// </summary>
		const string METRIC_SCIENCE_USE_RID = "Доля дохода от лицензионных соглашений от объема внебюджетных средств лаборатории в год";
		
		/// <summary>
		/// Перечень публикаций, написанных в ходе исследований, проводимых в лаборатории
		/// </summary>
		const string METRIC_SCIENCE_PUBLICATIONS = "Перечень публикаций, написанных в ходе исследований, проводимых в лаборатории";
		
		/// <summary>
		/// Перечень докладов на ведущих международных научных (научно-практических) конференциях в Российской Федерации и за рубежом
		/// </summary>
		const string METRIC_SCIENCE_EVENT_LIST_REPORT = "Перечень докладов на ведущих международных научных (научно-практических) конференциях в Российской Федерации и за рубежом";
		
		/// <summary>
		/// Перечень материалов докладов на ведущих международных научных (научно-практических) конференциях в Российской Федерации и за рубежом
		/// </summary>
		const string METRIC_SCIENCE_EVENT_LIST_MATERIAL = "Перечень материалов докладов на ведущих международных научных (научно-практических) конференциях в Российской Федерации и за рубежом";
		
		/// <summary>
		/// Перечень выставок (форум, салон, конгресс), в которых участвовали сотрудники лаборитории с презентацией разработок
		/// </summary>
		const string METRIC_SCIENCE_EVENT_LIST_EXHIBITIONS = "Перечень выставок (форум, салон, конгресс), в которых участвовали сотрудники лаборитории с презентацией разработок";
		
		/// <summary>
		/// Количество договоров на выполнение НИОКТР, в которых задействованы сотрудники лаборатории, в год
		/// </summary>
		const string METRIC_NIOKTR_COUNT_DOGOVOR = "Количество договоров на выполнение НИОКТР, в которых задействованы сотрудники лаборатории, в год";
		
		/// <summary>
		/// Доля студентов и аспирантов очной формы обучения (магистратуры и старших курсов бакалавриата и специалитета), трудоустроенных в лабораторию для участия в выполнении НИОКТР, в общем числе сотрудников лаборатории
		/// </summary>
		const string METRIC_NIOKTR_COUNT_STUDENTS = "Доля студентов и аспирантов очной формы обучения (магистратуры и старших курсов бакалавриата и специалитета), трудоустроенных в лабораторию для участия в выполнении НИОКТР, в общем числе сотрудников лаборатории";

		/// <summary>
		/// Перечень заказчиков НИОКТР, реализуемых (реализованных) в лаборатории
		/// </summary>
		const string METRIC_CUSTOMERS = "Перечень заказчиков НИОКТР, реализуемых (реализованных) в лаборатории";
		
		/// <summary>
		/// Описание результатов, полученных в ходе реализации НИОКТР
		/// </summary>
		const string METRIC_NIOKTR_RESULT = "Описание результатов, полученных в ходе реализации НИОКТР";
		#endregion
		#endregion
		#region Вспомогательные классы
		public class EmployeeDataScientists
		{
			/// <summary>
			/// ФИО физического лица
			/// </summary>
			public string FIO {
				get;
				set;
			}

			/// <summary>
			/// Ученые звания
			/// </summary>
			public List<DataScientists> AcademicTitles {
				get;
				set;
			}

			/// <summary>
			/// Ученые степени
			/// </summary>
			public List<DataScientists> DegreesTitles {
				get;
				set;
			}

			public EmployeeDataScientists ()
			{
				AcademicTitles = new List<DataScientists> ();
				DegreesTitles = new List<DataScientists> ();
			}
		}

		/// <summary>
		/// Данные об ученых степенях или званиях
		/// </summary>
		public class DataScientists
		{
			/// <summary>
			/// Название
			/// </summary>
			public string Name {
				get;
				set;
			}

			/// <summary>
			/// Дата
			/// </summary>
			public DateTime Data {
				get;
				set;
			}
		}

		public class DocumentPersonnel
		{
			/// <summary>
			/// Название
			/// </summary>
			public string Name {
				get;
				set;
			}

			/// <summary>
			/// Дата
			/// </summary>
			public DateTime Data {
				get;
				set;
			}
		}

		public class Personnel
		{
			/// <summary>
			/// ФИО сотрудника
			/// </summary>
			public string FIO {
				get;
				set;
			}

			/// <summary>
			/// Код сотрудника
			/// </summary>
			public string NumberSotrudnik {
				get;
				set;
			}

			/// <summary>
			/// Должность
			/// </summary>
			public string Position {
				get;
				set;
			}

			/// <summary>
			/// Подразделение
			/// </summary>
			public string Subdivision {
				get;
				set;
			}

			public List<DocumentPersonnel> Documents {
				get;
				set;
			}

			public Personnel ()
			{
				Documents = new List<DocumentPersonnel> ();
			}

			public bool IsMainWork {
				get;
				set;
			}
		}

		#endregion
		public void Test (Context context)
		{
			
		}
		
		#region Показатели
		/// <summary>
		/// Обновление данных по показателям
		/// </summary>
		/// <param name="context"></param>
		public void UpdateMetricsData (Context context)
		{
			var labs = PublicAPI.Objects.UserObjects.UserPodrazdeleniya.Find (FetchOptions.All);
			foreach (var lab in labs)
			{
				P_Pokazateli1S_Pokazateli item = new P_Pokazateli1S_Pokazateli();
				item.Laboratoriya = lab;
				context.Pokazateli.Add(item);
				
				UpdateMetricPersonal (context, lab);
				UpdateMetricEquipment (context, lab);
				UpdateMetricTechnologies (context, lab);
				UpdateMetricEducational (context, lab);
				UpdateMetricScientificActivity (context, lab);
				UpdateMetricCustomers (context, lab);
				UpdateMetricNIOKTR(context, lab);
			}
		}

		/// <summary>
		/// Получает или создает&получает метрику лаборатории
		/// </summary>
		/// <param name="lab">Лаборатория</param>
		/// <param name="nameMetric">Название показателя</param>
		/// <returns>Показатель лаборатории</returns>
		public EleWise.ELMA.ConfigurationModel.PokazateliLaboratorii GetOrCreateMetrick (EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab, string nameMetric)
		{
			var metric = PublicAPI.Objects.UserObjects.UserPokazatelj.Find (string.Format ("Naimenovanie = '{0}'", nameMetric)).FirstOrDefault ();
			if (metric == null) {
				throw new Exception (string.Format ("В справочнике 'Показатели' не найден показатель {0}", nameMetric));
			}
			var metricByLab = PublicAPI.Objects.UserObjects.UserPokazateliLaboratorii.Find (string.Format ("Podrazdelenie = {0} AND Pokazatelj = {1}", lab.Id, metric.Id)).SingleOrDefault ();
			if (metricByLab == null) {
				metricByLab = PublicAPI.Objects.UserObjects.UserPokazateliLaboratorii.Create ();
				metricByLab.Podrazdelenie = lab;
				metricByLab.Pokazatelj = metric;
				metricByLab.Save ();
			}
			return metricByLab;
		}

		/// <summary>
		/// Обновляет данные в справочнике история показателей
		/// </summary>
		/// <param name="lab">Лаборатория</param>
		/// <param name="metric">Показатель</param>
		/// <param name="val">Значение</param>
		public void UpdateHistoryMetric (EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab, EleWise.ELMA.ConfigurationModel.Pokazatelj metric, string val)
		{
			var history = PublicAPI.Objects.UserObjects.UserIstoriyaPokazateley.Find (string.Format ("Podrazdelenie = {0} AND Pokazatelj = {1}", lab.Id, metric.Id));
			if (history.Count == 0) {
				CreateHistoryMetrick (lab, metric, val);
				return;
			}
			var lasthistory = history.OrderBy (x => x.CreationDate).Last ();
			if (lasthistory.Znachenie != val) {
				CreateHistoryMetrick (lab, metric, val);
			}
		}

		/// <summary>
		/// Создает новую запись в справочнике история показателей
		/// </summary>
		/// <param name="lab">Лаборатория</param>
		/// <param name="metric">Показатель</param>
		/// <param name="val">Значение</param>
		public void CreateHistoryMetrick (EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab, EleWise.ELMA.ConfigurationModel.Pokazatelj metric, string val)
		{
			var history = PublicAPI.Objects.UserObjects.UserIstoriyaPokazateley.Create ();
			history.Podrazdelenie = lab;
			history.Pokazatelj = metric;
			history.Znachenie = val;
			history.Save ();
		}

		#region Кадровый состав лаборатории
		/// <summary>
		/// Обновление данных по показателям Кадровый состав лаборатории
		/// </summary>
		/// <param name="lab">Лаборатория</param>
		public void UpdateMetricPersonal (Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{
			UpdateSotrudnikCount (context, lab);
			UpdateSotrudnikTravelCount (context, lab);
			UpdateShareSotrudnikBefore39 (context, lab);
			UpdateShareSotrudnikCandidateCount (context, lab);
			UpdateShareSotrudnikDoctorCount (context, lab);
			UpdateShareSotrudnikNPR (context, lab);
			UpdateListCandidateWork (context, lab);
			UpdateListDoctorWork (context, lab);
		}

		/// <summary>
		/// Обновление показатея 'Количество сотрудников лаборатории'
		/// </summary>
		/// <param name="lab">Лаборатория</param>
		public void UpdateSotrudnikCount (Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{
			var metric = GetOrCreateMetrick (lab, METRIC_COUNT_SOTRUDNIK);
			var sotrudnikCount = PublicAPI.Objects.UserObjects.UserKadrovayaIstoriyaSotrudnika.Find (string.Format ("Podrazdelenie = {0}", lab.Id)).Where (x => x.DataUvoljneniya == null).Count ();
			metric.Znachenie = sotrudnikCount.ToString ();
			context.Pokazateli.Last().PokazateliLaboratorii.Add(metric);
			UpdateHistoryMetric (lab, metric.Pokazatelj, metric.Znachenie);
		}

		/// <summary>
		/// Обновление показатея 'Количество исследователей, направленных на работу (стажировку) в зарубежные организации (научные организации/университеты)'
		/// </summary>
		/// <param name="lab">Лаборатория</param>
		public void UpdateSotrudnikTravelCount (Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{
			var metric = GetOrCreateMetrick (lab, METRIC_COUNT_TRAVELERS);
			var travelCount = PublicAPI.Objects.UserObjects.UserNapravlenieNaRabotuStazhirovkuVZarubezhnyeOrganizaciiNauchnyeOrganizaciiUniversitety.Find (string.Format ("DataNachalaStazhirovkiRaboty <= DateTime('Now') AND DataOkonchaniyaStazhirovkiRaboty >= DateTime('Now') AND Podrazdelenie = {0}", lab.Id));
			metric.Znachenie = travelCount.Count.ToString ();
			context.Pokazateli.Last().PokazateliLaboratorii.Add(metric);
		}

		/// <summary>
		/// Обновление показатея 'Доля исследователей в составе лаборатории в возрасте до 39 лет'
		/// </summary>
		/// <param name="lab">Лаборатория</param>		
		public void UpdateShareSotrudnikBefore39 (Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{

			var metric = GetOrCreateMetrick (lab, METRIC_SHARE_SOTRUDNIK_BEFORE_39);
			var workSotrudniks = PublicAPI.Objects.UserObjects.UserKadrovayaIstoriyaSotrudnika.Find (string.Format ("Podrazdelenie = {0}", lab.Id)).Where (x => x.DataUvoljneniya == null);
			var sotrudnikCount = workSotrudniks.Count ();
			var skResearcher = workSotrudniks.Where (x => x.Dolzhnostj.ToLower ().Contains ("научный сотрудник") && x.OsnovnoeMestoRaboty);
			var researchCount = 0;
			foreach (var data in skResearcher) {
				if (data.Sotrudnik.Ssylka1sRD == null)
					continue;
				if (data.Sotrudnik.Ssylka1sRD.FizicheskoeLico.DataRozhdeniya == null)
					continue;
				var dataBirthday = data.Sotrudnik.Ssylka1sRD.FizicheskoeLico.DataRozhdeniya;
				var today = DateTime.Today;
				// Calculate the age.
				var age = today.Year - dataBirthday.Year;
				// Go back to the year in which the person was born in case of a leap year
				if (dataBirthday.Date > today.AddYears (-age))
					age--;
				if (age < 39) {
					researchCount++;
				}
			}
			double val = (((double)researchCount / (double)sotrudnikCount) * 100);
			if (Double.IsNaN (val)) {
				val = 0;
			}
			metric.Znachenie = val.ToString ();
			context.Pokazateli.Last().PokazateliLaboratorii.Add(metric);
			UpdateHistoryMetric (lab, metric.Pokazatelj, metric.Znachenie);
		}

		/// <summary>
		/// Обновление показатея 'Доля сотрудников лаборатории с наличием степени кандидата наук'
		/// </summary>
		/// <param name="lab">Лаборатория</param>		
		public void UpdateShareSotrudnikCandidateCount (Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{
			var metric = GetOrCreateMetrick (lab, METRIC_SHARE_SOTRUDNIK_CANDIDATE);
			var workSotrudniks = PublicAPI.Objects.UserObjects.UserKadrovayaIstoriyaSotrudnika.Find (string.Format ("Podrazdelenie = {0}", lab.Id)).Where (x => x.DataUvoljneniya == null);
			var candidate = workSotrudniks.Where (x => !string.IsNullOrEmpty (x.Sotrudnik.UchenoeZvanie) && x.Sotrudnik.UchenoeZvanie.ToLower ().Contains ("кандидат"));
			double val = (((double)candidate.Count () / (double)workSotrudniks.Count ()) * 100);
			if (Double.IsNaN (val)) {
				val = 0;
			}
			metric.Znachenie = val.ToString ();
			context.Pokazateli.Last().PokazateliLaboratorii.Add(metric);
			UpdateHistoryMetric (lab, metric.Pokazatelj, metric.Znachenie);
		}

		/// <summary>
		/// Обновление показатея 'Доля сотрудников лаборатории с наличием степени доктора наук'
		/// </summary>
		/// <param name="lab">Лаборатория</param>		
		public void UpdateShareSotrudnikDoctorCount (Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{
			var metric = GetOrCreateMetrick (lab, METRIC_SHARE_SOTRUDNIK_DOCTOR);
			var workSotrudniks = PublicAPI.Objects.UserObjects.UserKadrovayaIstoriyaSotrudnika.Find (string.Format ("Podrazdelenie = {0}", lab.Id)).Where (x => x.DataUvoljneniya == null);
			var doctor = workSotrudniks.Where (x => !string.IsNullOrEmpty (x.Sotrudnik.UchenoeZvanie) && x.Sotrudnik.UchenoeZvanie.ToLower ().Contains ("доктор"));
			double val = (((double)doctor.Count () / (double)workSotrudniks.Count ()) * 100);
			if (Double.IsNaN (val)) {
				val = 0;
			}

			metric.Znachenie = val.ToString ();
			context.Pokazateli.Last().PokazateliLaboratorii.Add(metric);
			UpdateHistoryMetric (lab, metric.Pokazatelj, metric.Znachenie);
		}

		/// <summary>
		/// Обновление показатея 'Доля НПР кафедр, трудоустроенных в лабораторию, в общей численности лаборатории'
		/// </summary>
		/// <param name="lab">Лаборатория</param>		
		public void UpdateShareSotrudnikNPR (Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{
			var metric = GetOrCreateMetrick (lab, METRIC_SHARE_SOTRUDNIK_NPR_DEPARTMENT);

			var workSotrudnikslab = PublicAPI.Objects.UserObjects.UserKadrovayaIstoriyaSotrudnika.Find (string.Format ("Podrazdelenie = {0}", lab.Id)).Where (x => x.DataUvoljneniya == null);
			var workSotrudniksAll = PublicAPI.Objects.UserObjects.UserKadrovayaIstoriyaSotrudnika.Find (FetchOptions.All).Where (x => x.DataUvoljneniya == null);
			
			long countUserNpf = 0;
			foreach(var element in workSotrudnikslab)
			{
				var sk = workSotrudniksAll.Where(x => x.Sotrudnik.Id == element.Sotrudnik.Id);
				if(sk.Count() > 1)
				{
					foreach(var itemSk in sk)
					{
						if (itemSk.Podrazdelenie.Id == lab.Id)
							continue;
							
						if(itemSk.Podrazdelenie.PolnoeNaimenovanie.ToLower().Contains("кафедр"))
						{
							countUserNpf++;
							break;
						}
					}
				}
			}
			
			
			
			double val = (((double)countUserNpf / (double)workSotrudnikslab.Count ()) * 100);
			if (Double.IsNaN (val)) {
				val = 0;
			}

			metric.Znachenie = val.ToString ();
			context.Pokazateli.Last().PokazateliLaboratorii.Add(metric);
			UpdateHistoryMetric (lab, metric.Pokazatelj, metric.Znachenie);
			
		}

		/// <summary>
		/// Обновление показатея 'Перечень защищенных диссертаций сотрудниками лаборатории (кандидатские)'
		/// </summary>
		/// <param name="lab">Лаборатория</param>		
		public void UpdateListCandidateWork (Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{
			var metric = GetOrCreateMetrick (lab, METRIC_LIST_CANDIDATE_WORK);
			
			var workSk = PublicAPI.Objects.UserObjects.UserKadrovayaIstoriyaSotrudnika.Find(string.Format("Podrazdelenie = {0} AND DataUvoljneniya is NULL", lab.Id));
			var currentActs = PublicAPI.Objects.UserObjects.UserPokazateliLaboratorii.Find(string.Format("Podrazdelenie = {0} AND Naimenovanie = '{1}'",lab.Id,METRIC_LIST_CANDIDATE_WORK));
			List<string> Dissertation = new List<string>();
			
			if(currentActs.Count > 0)
			{
				var row = currentActs.Single();
				var data = row.Znachenie.Split(';');
				if(data.Length > 0)
				{
					foreach(var item in data)
					{
						Dissertation.Add(item);
					}
				}
			}
			
			foreach(var sk in workSk)
			{
				var dissert = PublicAPI.Objects.UserObjects.UserDissertacii.Find(string.Format("Avtory = {0} AND TipDissertacii = DropDownItem('Кандидатская')",sk.Sotrudnik.Id));
				if(dissert.Count > 0)
				{
					foreach(var element in dissert)
					{
						if(!Dissertation.Any(x => x == element.Nazvanie))
						{
							Dissertation.Add(element.Nazvanie);
						}
					}
				}
			}
			
			
			var result = string.Join(";", Dissertation);
			metric.Znachenie = result;
			context.Pokazateli.Last().PokazateliLaboratorii.Add(metric);
			UpdateHistoryMetric (lab, metric.Pokazatelj, metric.Znachenie);
		}

		/// <summary>
		/// Обновление показатея 'Перечень защищенных диссертаций сотрудниками лаборатории (докторские)'
		/// </summary>
		/// <param name="lab">Лаборатория</param>		
		public void UpdateListDoctorWork (Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{
			var metric = GetOrCreateMetrick (lab, METRIC_LIST_DOCTORAL_WORK);
			
			var workSk = PublicAPI.Objects.UserObjects.UserKadrovayaIstoriyaSotrudnika.Find(string.Format("Podrazdelenie = {0} AND DataUvoljneniya is NULL", lab.Id));
			var currentActs = PublicAPI.Objects.UserObjects.UserPokazateliLaboratorii.Find(string.Format("Podrazdelenie = {0} AND Naimenovanie = '{1}'",lab.Id,METRIC_LIST_DOCTORAL_WORK));
			List<string> Dissertation = new List<string>();
			
			if(currentActs.Count > 0)
			{
				var row = currentActs.Single();
				var data = row.Znachenie.Split(';');
				if(data.Length > 0)
				{
					foreach(var item in data)
					{
						Dissertation.Add(item);
					}
				}
			}
			
			foreach(var sk in workSk)
			{
				var dissert = PublicAPI.Objects.UserObjects.UserDissertacii.Find(string.Format("Avtory = {0} AND TipDissertacii = DropDownItem('Докторская')",sk.Sotrudnik.Id));
				if(dissert.Count > 0)
				{
					foreach(var element in dissert)
					{
						if(!Dissertation.Any(x => x == element.Nazvanie))
						{
							Dissertation.Add(element.Nazvanie);
						}
					}
				}
			}
			
			
			var result = string.Join(";", Dissertation);
			metric.Znachenie = result;
			context.Pokazateli.Last().PokazateliLaboratorii.Add(metric);
			UpdateHistoryMetric (lab, metric.Pokazatelj, metric.Znachenie);
		}

		#endregion
		
		
		
		
		#region Показатели Оборудование
		public void UpdateMetricEquipment (Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{
			var metric = GetOrCreateMetrick (lab, METRIC_EQUIPMENT_LIST_EQUIPMENT);
			string result = "";
			var equipment = lab.IspoljzuemoeOborudovanie.ToList().OrderBy(x => x.Naimenovanie);
			foreach(var element in equipment)
			{
				result += element.Naimenovanie + ";" + Environment.NewLine;
			}
			metric.Znachenie = result;
			context.Pokazateli.Last().PokazateliLaboratorii.Add(metric);
			UpdateHistoryMetric (lab, metric.Pokazatelj, metric.Znachenie);
		}
		#endregion
		
		#region Показатели Технологии
		public void UpdateMetricTechnologies (Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{
			var metric = GetOrCreateMetrick (lab, METRIC_TECHNOLOGY_LIST_TECHNOLOGY);
			string result = "";
			var tehnology = lab.IspoljzuemyeTehnologii.ToList().OrderBy(x => x.Naimenovanie);
			foreach(var element in tehnology)
			{
				result += element.Naimenovanie + ";" + Environment.NewLine;
			}
			metric.Znachenie = result;
			context.Pokazateli.Last().PokazateliLaboratorii.Add(metric);
			UpdateHistoryMetric (lab, metric.Pokazatelj, metric.Znachenie);
		}
		#endregion
		
		#region Показатели Образовательная деятельность лаборатории
		public void UpdateMetricEducational (Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{
			UpdateListStudents(context, lab);
			UpdateListOfModule(context, lab);
		}
		
		/// <summary>
		/// Показатель Количество обучающихся (студенты СПО, студенты ВО), прошедших (проходящих) практику в лаборатории
		/// </summary>
		public void UpdateListStudents(Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{
			var metric = GetOrCreateMetrick (lab, METRIC_STUDY_LIST_STUDENTS);
			var students = PublicAPI.Objects.UserObjects.UserZhurnalPraktik.Find(FetchOptions.All).Count(x => x.Podrazdelenie1C != null && x.Podrazdelenie1C.Kod == lab.Ssylka1sRD.Kod);
			metric.Znachenie = students.ToString();
			context.Pokazateli.Last().PokazateliLaboratorii.Add(metric);
			UpdateHistoryMetric (lab, metric.Pokazatelj, metric.Znachenie);
		}
		
		/// <summary>
		/// Показатель Перечень образовательных модулей, которые ведут сотрудники лаборатории
		/// </summary>
		public void UpdateListOfModule(Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{
			var metric = GetOrCreateMetrick (lab, METRIC_STUDY_LIST_STUDY_MODULES);
			var modules = PublicAPI.Objects.UserObjects.UserObrazovateljnyeModuli.Find(string.Format("Podrazdelenie = {0}", lab.Id)).ToList().OrderBy(x => x.Naimenovanie);
			string result = "";
			foreach(var element in modules)
			{
				result += element.Naimenovanie + ";" + Environment.NewLine;
			}
			metric.Znachenie = result.ToString();
			context.Pokazateli.Last().PokazateliLaboratorii.Add(metric);
			UpdateHistoryMetric (lab, metric.Pokazatelj, metric.Znachenie);
		}
		#endregion
		
		#region Показатели Научная деятельность лаборатории
		public void UpdateMetricScientificActivity (Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{
			UpdateRID(context, lab);
			UpdateExecuteRID(context, lab);
			UpdatePublications(context, lab);
			UpdateParticipation(context, lab);
			UpdateHoldingNIOKTR(context, lab);
			UpdateStudentInNIOKTR(context, lab);
		}
		
		/// <summary>
		/// Показатель Список патентов на изобретение, полезную модель, полученных в ходе работ, проводимых в лаборатории / Список свидетельств о гос.регистрации программы для ЭВМ, баз данных, полученных в ходе работ, проводимых в лаборатории
		/// </summary>
		public void UpdateRID(Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{
			var patent = GetOrCreateMetrick (lab, METRIC_SCIENCE_RID_PATTENT);
			patent.Znachenie = "";
			UpdateHistoryMetric (lab, patent.Pokazatelj, patent.Znachenie);
			
			var metric = GetOrCreateMetrick (lab, METRIC_SCIENCE_RID_SERTIFICATE);
			metric.Znachenie = "";
			context.Pokazateli.Last().PokazateliLaboratorii.Add(metric);
			UpdateHistoryMetric (lab, metric.Pokazatelj, metric.Znachenie);
		}
		
		/// <summary>
		/// Показатель Доля дохода от лицензионных соглашений от объема внебюджетных средств лаборатории в год
		/// </summary>
		public void UpdateExecuteRID(Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{
			var metric = GetOrCreateMetrick (lab, METRIC_SCIENCE_USE_RID);
			metric.Znachenie = "0";
			context.Pokazateli.Last().PokazateliLaboratorii.Add(metric);
			UpdateHistoryMetric (lab, metric.Pokazatelj, metric.Znachenie);
		}
		
		/// <summary>
		/// Показатель Перечень публикаций, написанных в ходе исследований, проводимых в лаборатории
		/// </summary>
		public void UpdatePublications(Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{
			var metric = GetOrCreateMetrick (lab, METRIC_SCIENCE_PUBLICATIONS);
			metric.Znachenie = "";
			context.Pokazateli.Last().PokazateliLaboratorii.Add(metric);
			UpdateHistoryMetric (lab, metric.Pokazatelj, metric.Znachenie);
		}
		
		
		/// <summary>
		/// Участие в мероприятиях, конференциях, семинарах
		/// </summary>
		public void UpdateParticipation(Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{
			UpdateListDokladRF(context, lab);
			UpdateListMaterailRF(context, lab);
			UpdateListForums(context, lab);
		}
		
		/// <summary>
		/// Перечень докладов на ведущих международных научных (научно-практических) конференциях в Российской Федерации и за рубежом
		/// </summary>
		/// <param name="Context"></param>
		/// <param name="lab"></param>
		public void UpdateListDokladRF(Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{
			var metric = GetOrCreateMetrick (lab, METRIC_SCIENCE_EVENT_LIST_REPORT);
			metric.Znachenie = "";
			context.Pokazateli.Last().PokazateliLaboratorii.Add(metric);
			UpdateHistoryMetric (lab, metric.Pokazatelj, metric.Znachenie);
		}
		
		/// <summary>
		/// Перечень материалов докладов на ведущих международных научных (научно-практических) конференциях в Российской Федерации и за рубежом
		/// </summary>
		/// <param name="Context"></param>
		/// <param name="lab"></param>
		public void UpdateListMaterailRF(Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{
			var metric = GetOrCreateMetrick (lab, METRIC_SCIENCE_EVENT_LIST_MATERIAL);
			metric.Znachenie = "";
			context.Pokazateli.Last().PokazateliLaboratorii.Add(metric);
			UpdateHistoryMetric (lab, metric.Pokazatelj, metric.Znachenie);
		}
		
		/// <summary>
		/// Перечень выставок (форум, салон, конгресс), в которых участвовали сотрудники лаборитории с презентацией разработок
		/// </summary>
		/// <param name="Context"></param>
		/// <param name="lab"></param>
		public void UpdateListForums(Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{
			var metric = GetOrCreateMetrick (lab, METRIC_SCIENCE_EVENT_LIST_EXHIBITIONS);
			metric.Znachenie = "";
			context.Pokazateli.Last().PokazateliLaboratorii.Add(metric);
			UpdateHistoryMetric (lab, metric.Pokazatelj, metric.Znachenie);
		}
		
		
		/// <summary>
		/// Показатель Количество договоров на выполнение НИОКТР, в которых задействованы сотрудники лаборатории, в год
		/// </summary>
		public void UpdateHoldingNIOKTR(Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{
			var metric = GetOrCreateMetrick (lab, METRIC_NIOKTR_COUNT_DOGOVOR);
			metric.Znachenie = "0";
			context.Pokazateli.Last().PokazateliLaboratorii.Add(metric);
			UpdateHistoryMetric (lab, metric.Pokazatelj, metric.Znachenie);
		}
		
		/// <summary>
		/// Показатель Доля студентов и аспирантов очной формы обучения (магистратуры и старших курсов бакалавриата и специалитета), трудоустроенных в лабораторию для участия в выполнении НИОКТР, в общем числе сотрудников лаборатории
		/// </summary>
		public void UpdateStudentInNIOKTR(Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{
			var metric = GetOrCreateMetrick (lab, METRIC_NIOKTR_COUNT_STUDENTS);
			metric.Znachenie = "0";
			context.Pokazateli.Last().PokazateliLaboratorii.Add(metric);
			UpdateHistoryMetric (lab, metric.Pokazatelj, metric.Znachenie);
		}
		
		
		#endregion
		
		#region Показатели Заказчики
		public void UpdateMetricCustomers (Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{
			var metric = GetOrCreateMetrick (lab, METRIC_CUSTOMERS);
			metric.Znachenie = "";
			context.Pokazateli.Last().PokazateliLaboratorii.Add(metric);
			UpdateHistoryMetric (lab, metric.Pokazatelj, metric.Znachenie);
		}
		#endregion
		
		#region Показатели NIOKTR
		public void UpdateMetricNIOKTR (Context context, EleWise.ELMA.ConfigurationModel.Podrazdeleniya lab)
		{
			var metric = GetOrCreateMetrick (lab, METRIC_NIOKTR_RESULT);
			metric.Znachenie = "";
			context.Pokazateli.Last().PokazateliLaboratorii.Add(metric);
			UpdateHistoryMetric (lab, metric.Pokazatelj, metric.Znachenie);
		}
		#endregion
		
		
		#endregion
		/// <summary>
		/// NoExistProcess
		/// </summary>
		/// <param name="context">Контекст процесса</param>
		/// <param name="GatewayVar"></param>
		public virtual bool NoExistProcess (Context context, object GatewayVar)
		{
			var processes = PublicAPI.Processes.WorkflowInstance.Filter ().ProcessHeaderId(PROCESS_HEADER_ID).GeneralStatus (PublicAPI.Enums.Workflow.WorkflowInstanceGeneralStatus.Current).Find ();
			return !processes.Any (x => x.Id != context.WorkflowInstance.Id);
		}

		/// <summary>
		/// ExistProcess
		/// </summary>
		/// <param name="context">Контекст процесса</param>
		/// <param name="GatewayVar"></param>
		public virtual bool ExistProcess (Context context, object GatewayVar)
		{
			var processes = PublicAPI.Processes.WorkflowInstance.Filter ().ProcessHeaderId(PROCESS_HEADER_ID).GeneralStatus (PublicAPI.Enums.Workflow.WorkflowInstanceGeneralStatus.Current).Find ();
			return processes.Any (x => x.Id != context.WorkflowInstance.Id);
		}
	}

}
