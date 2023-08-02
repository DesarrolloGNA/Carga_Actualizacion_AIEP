using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Text;

namespace Carga_Actualizacion_AIEP
{
    class Program
    {

        public static List<ACTUALIZACIONES> Lista_Actualizaciones = new List<ACTUALIZACIONES>();
        
        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Encoding encoding = Encoding.GetEncoding(1252);

            DirectoryInfo di = new DirectoryInfo(@"C:\aiep\Actualizaciones\");
            FileInfo[] files = di.GetFiles("*.xlsx");


            foreach (FileInfo file in files)
            {
                String Ruta_Archivo = @"C:\aiep\Actualizaciones\" + file.Name;

                Console.WriteLine("Lectura Archivo :" + file.Name.ToString());
                /*-------------------------------------------------*/
                /*             LECTURA DE EXCEL                    */
                /*-------------------------------------------------*/
                Leer_Excel(Ruta_Archivo);


                /*-------------------------------------------------*/
                /*             CARGA A BASE DE DATOS               */
                /*-------------------------------------------------*/

                int cantidad = 0;
                foreach (var i in Lista_Actualizaciones)
                {
                    cantidad++;
                    Console.WriteLine("Insertando Registro :" + cantidad.ToString());

                    /*-------------------------------------------------*/
                    /*            CARGA ALA BASE DE DATOS              */
                    /*-------------------------------------------------*/

                    string connstring = @"Data Source=192.168.0.5; Initial Catalog=EJFDES; Persist Security Info=True; User ID=sa; Password=w2003ejf103;";
                    using (SqlConnection con = new SqlConnection(connstring))
                    {
                        con.Open();

                        string commandString = @"INSERT INTO EJFDES.dbo.Historico_Actualizacion_AIEP (FECHA_CARGA,ID_MANDANTE,RUT_ACEPTANTE,RUT_ALUMNO,NRO_DOCUMENTO
                        ,SALDO_DOCUMENTO,FECHA_VENCIMIENTO)VALUES(GETDATE(),118,@RUT_ACEPTANTE,@RUT_ALUMNO,@NRO_DOCUMENTO,@SALDO_DOCUMENTO,@FECHA_VENCIMIENTO)";

                        SqlCommand cmd = new SqlCommand(commandString, con);
                        cmd.Parameters.AddWithValue("@RUT_ACEPTANTE", i.RUT_ACEPTANTE);
                        cmd.Parameters.AddWithValue("@RUT_ALUMNO", i.RUT_ALUMNO);
                        cmd.Parameters.AddWithValue("@NRO_DOCUMENTO", i.NRO_DOCUMENTO);
                        cmd.Parameters.AddWithValue("@SALDO_DOCUMENTO", i.SALDO_DOCUMENTO);
                        cmd.Parameters.AddWithValue("@FECHA_VENCIMIENTO", i.FECHA_VENCIMIENTO);

                        cmd.ExecuteNonQuery();
                        con.Close();
                    }
                }

            }

        }

        public static void Leer_Excel(string ruta)
        {
            int contador = 1;
            DateTime Fecha_Actual = DateTime.Now;
            //===========================================================
            // LECTURA DEL ARCHIVO                                              
            //===========================================================
            using (var stream = File.Open(ruta, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration()
                {
                    FallbackEncoding = Encoding.UTF8,
                    LeaveOpen = false,
                    AnalyzeInitialCsvRows = 0,
                }))
                {
                    
                    do
                    {
                        //recorro el excel 
                        while (reader.Read())
                        {
                            Console.WriteLine("Leyendo Registro :" + contador.ToString());

                            //omito la cabecera
                            if (contador > 1)
                                {
                                        try
                                        {
                                            ACTUALIZACIONES Input = new ACTUALIZACIONES();
                                            Input.RUT_ACEPTANTE = reader.GetValue(2) != null ? reader.GetValue(2).ToString() : "";
                                            Input.RUT_ALUMNO = reader.GetValue(3) != null ? reader.GetValue(3).ToString() : "";
                                            Input.NRO_DOCUMENTO = reader.GetValue(6) != null ? reader.GetValue(6).ToString() : "";
                                            Input.SALDO_DOCUMENTO = reader.GetValue(9) != null ? int.Parse(reader.GetValue(9).ToString()) : 0;
                                            Input.FECHA_VENCIMIENTO = reader.GetValue(7) != null ? DateTime.Parse(reader.GetValue(7).ToString()) : DateTime.Parse("19000101");
                                            Lista_Actualizaciones.Add(Input);
                                        }
                                        catch (Exception Ex)
                                        {
                                            Console.WriteLine("Error :"+ Ex.Message.ToString() + ", En registro :" + contador.ToString());
                                        }

                                }
                                contador++;                         
                        }
                    } while (reader.NextResult());
                }
            }
        }

        public class ACTUALIZACIONES
        {
            public string RUT_ACEPTANTE { get; set; }
            public string RUT_ALUMNO { get; set; }
            public string NRO_DOCUMENTO { get; set; }
            public int SALDO_DOCUMENTO { get; set; }
            public DateTime FECHA_VENCIMIENTO { get; set; }

        }

    }
}
