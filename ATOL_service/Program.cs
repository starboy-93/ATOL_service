using Atol.Drivers10.Fptr;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace ATOL_service
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.InputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine("Синхронизатор времени на ФР.\n");
            try
            {
                var fptr = new Fptr();
                var connection = fptr.open(); //подключаемся к ККТ
                ConnectionSuccessfull(fptr, connection);
            }
            catch (System.IO.FileNotFoundException)
            {
                Console.WriteLine("Атол \"Тест драйвера ККТ\" не установлен.");
                Console.ReadLine();
            }
        }

        private static void ConnectionSuccessfull(Fptr fptr, int connection)
        {
            if (connection != -1) //если драйвер не занят
            {
                String version = fptr.version();
                Console.WriteLine($"Версия драйвера АТОЛ:  {version}");

                //проверяем открыта ли смена
                fptr.setParam(Constants.LIBFPTR_PARAM_DATA_TYPE, Constants.LIBFPTR_DT_SHIFT_STATE);
                fptr.queryData();
                uint state = fptr.getParamInt(Constants.LIBFPTR_PARAM_SHIFT_STATE);
                if (state == 0) //если смена открыта
                {
                    Console.WriteLine("Смена закрыта. Время будет синхронизировано с реальным.\n");
                    //получаем текущее дату и время с ФР
                    fptr.setParam(Constants.LIBFPTR_PARAM_DATA_TYPE, Constants.LIBFPTR_DT_DATE_TIME);
                    fptr.queryData();
                    DateTime dateTime = fptr.getParamDateTime(Constants.LIBFPTR_PARAM_DATE_TIME);
                    Console.WriteLine("Время на ФР было: " + dateTime);

                    //получаем текущее дату и время с ПК
                    DateTime currentDateTime = DateTime.Now;

                    //задаём на ФР текущие время и дату, записываем
                    fptr.setParam(Constants.LIBFPTR_PARAM_DATE_TIME, currentDateTime);
                    fptr.writeDateTime();

                    //еще раз получаем время с ФР, выводим на экран, проверяем что всё отработало корректно
                    fptr.setParam(Constants.LIBFPTR_PARAM_DATA_TYPE, Constants.LIBFPTR_DT_DATE_TIME);
                    fptr.queryData();
                    dateTime = fptr.getParamDateTime(Constants.LIBFPTR_PARAM_DATE_TIME);
                    Console.WriteLine("Время на ФР стало: " + dateTime);
                    Console.WriteLine("Реальное время: " + DateTime.Now);
                }
                else //если смена открыта
                    Console.WriteLine("Смена открыта. Время выровнено не будет.\n");

                Console.WriteLine("-----------------------------------------------------------");
                OFDcheck(fptr);
                Console.WriteLine("-----------------------------------------------------------");
                FNckeck(fptr);

                fptr.destroy(); //отключаемся от ККТ
                Console.ReadLine(); //чтобы консоль не закрывалась 
            }
            else //если драйвер занят другой программой
            {
                FRisBusy(fptr);
            }
        }

        private static void FRisBusy(Fptr fptr)
        {
            Console.WriteLine("Запущена 1С или \"Тест Драйвера ККТ\". Для синхронизации времени их необходимо закрыть.\nХотите это сделать?");
            Console.WriteLine("1. Да. ВНИМАНИЕ, будут закрыты ВСЕ ОКНА 1C!!!\n2. Выйти");
            //var answer = Console.ReadKey();
            ConsoleKeyInfo keyInfo = Console.ReadKey();
            Console.WriteLine();
            if (keyInfo.KeyChar == '1')
            {
                AppKill("fptr10_t");
                AppKill("1cv8");
                fptr.destroy(); //отключаемся от ККТ
                var fptr2 = new Fptr();
                var connection2 = fptr2.open(); //подключаемся к ККТ заново
                ConnectionSuccessfull(fptr2, connection2);
            }
            else if (keyInfo.KeyChar == '2')
            {
                Console.WriteLine("Приложение будет закрыто");
                Environment.Exit(0);
            }
            else
            {
                Console.WriteLine("Неверный ввод. Попробуйте еще раз.");
                FRisBusy(fptr);
            }
        }


        private static void FNckeck(Fptr fptr)
        {

            //проверяем когда закончится ФН
            fptr.setParam(Constants.LIBFPTR_PARAM_FN_DATA_TYPE, Constants.LIBFPTR_FNDT_VALIDITY);
            fptr.fnQueryData();
            DateTime FNexpiryDateTime = fptr.getParamDateTime(Constants.LIBFPTR_PARAM_DATE_TIME);
            if (FNexpiryDateTime != new DateTime(1970, 1, 1, 0, 0, 0))
            {
                int daysBetween = (FNexpiryDateTime - DateTime.Now).Days;
                Console.WriteLine($"ФН на кассе заканчивается {FNexpiryDateTime.Date} через {daysBetween} дней");
                Console.WriteLine("Если до конца ФН меньше 7 дней, обратитесь к фиксикам");
            }
        }

        private static void OFDcheck(Fptr fptr)
        {
            //Запрос даты и времени последней успешной отправки документа в ОФД
            fptr.setParam(Constants.LIBFPTR_PARAM_DATA_TYPE, Constants.LIBFPTR_DT_LAST_SENT_OFD_DOCUMENT_DATE_TIME);
            fptr.queryData();

            DateTime OfdLastSuccessedDocTime = fptr.getParamDateTime(Constants.LIBFPTR_PARAM_DATE_TIME);
            if (OfdLastSuccessedDocTime == new DateTime(1970, 1, 1, 0, 0, 0))
                Console.WriteLine($"Дата и время последней успешной отправки документа в ОФД: <Не определено>");
            else
                Console.WriteLine($"Дата и время последней успешной отправки документа в ОФД: {OfdLastSuccessedDocTime}");
            Console.WriteLine("Если дата меньше вчерашней или не определена, обратитесь к фиксикам");
        }

        private static void AppKill(string proccess)
        {
            string processName = proccess; // Имя процесса без .exe

            // Получаем все процессы с заданным именем
            Process[] processes = Process.GetProcessesByName(processName);

            // Проверяем, есть ли такие процессы
            if (processes.Length > 0)
            {
                foreach (Process process in processes)
                {
                    try
                    {
                        process.Kill(); // Завершаем процесс
                        Console.WriteLine($"Процесс {process.ProcessName} (ID: {process.Id}) завершен.\n");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Не удалось завершить процесс {process.ProcessName} (ID: {process.Id}): {ex.Message} \n");
                    }
                }
            }
            else
            {
                Console.WriteLine($"Процессы с именем {processName} не найдены.\n");
            }
        }
    }
}

