using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataGridView
{
    public class ITextEvents : PdfPageEventHelper
    {

        // This is the contentbyte object of the writer
        PdfContentByte cb;

        // we will put the final number of pages in a template
        PdfTemplate footerTemplate;

        // this is the BaseFont we are going to use for the footer
        BaseFont bf = null;

        // This keeps track of the creation time
        DateTime PrintTime = DateTime.Now;


        public override void OnOpenDocument(PdfWriter writer, Document document)
        {
            try
            {
                bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb = writer.DirectContent;
                footerTemplate = cb.CreateTemplate(50, 50);
            }
            catch (DocumentException de)
            {
                //handle exception here
            }
            catch (System.IO.IOException ioe)
            {
                //handle exception here
            }
        }

        public override void OnEndPage(iTextSharp.text.pdf.PdfWriter writer, iTextSharp.text.Document document)
        {
            base.OnEndPage(writer, document);

            //Create PdfTable object
            //PdfPTable pdfTab = new PdfPTable(3);
            String text = "Pág: " + writer.PageNumber;

            // Creamos tabla para el título del informe
            PdfPTable pdfTable = new PdfPTable(3);
            //pdfTable.SpacingBefore = 5f;
            //pdfTable.WidthPercentage = 100;
            pdfTable.TotalWidth = document.PageSize.Width -40;
            pdfTable.DefaultCell.HorizontalAlignment = Element.ALIGN_LEFT;
            pdfTable.DefaultCell.VerticalAlignment = Element.ALIGN_CENTER;
            pdfTable.DefaultCell.BorderWidth = 0;
            pdfTable.DefaultCell.BorderWidthTop = 1;

            // Generamos los datos textos que van en el footer
            // Emision
            CultureInfo culture = new CultureInfo("pt-BR");
            Chunk c1 = new Chunk("Emisión: " + PrintTime.ToString("dd/MM/yyyy HH:mm:ss", culture), FontFactory.GetFont("Microsoft Sans Serif", 8));
            c1.Font.Color = new iTextSharp.text.BaseColor(0, 0, 0);

            //Titulo
            Chunk c2 = new Chunk("Programa Operativo Anual " + PrintTime.Year, FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD));
            c2.Font.Color = new iTextSharp.text.BaseColor(0, 0, 0);

            // Paginado
            Chunk c3 = new Chunk(text, FontFactory.GetFont("Microsoft Sans Serif", 8));
            c3.Font.Color = new iTextSharp.text.BaseColor(0, 0, 0);


            Phrase p1 = new Phrase();
            p1.Add(c1);
            pdfTable.AddCell(p1);
            pdfTable.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;
            Phrase p2 = new Phrase();
            p2.Add(c2);
            pdfTable.AddCell(p2);
            pdfTable.DefaultCell.HorizontalAlignment = Element.ALIGN_RIGHT;
            Phrase p3 = new Phrase();
            p3.Add(c3);
            pdfTable.AddCell(p3);

            //pdfTable.WriteSelectedRows(0, -1, 40, document.PageSize.Height - 30, writer.DirectContent);
            pdfTable.WriteSelectedRows(0, -1, 20, 30, writer.DirectContent);
            //pdfTable.WriteSelectedRows(0, -1, 40, 500, writer.DirectContent);


            /*
            // Agregamos fecha de emisión al Footer
            cb.BeginText();
            cb.SetFontAndSize(bf, 8);
            cb.SetTextMatrix(document.PageSize.GetRight(586), document.PageSize.GetBottom(40));
            CultureInfo culture = new CultureInfo("pt-BR");
            cb.ShowText("Emisión: " + PrintTime.ToString("dd/MM/yyyy HH:mm:ss", culture));
            cb.EndText();
            cb.AddTemplate(footerTemplate, document.PageSize.GetRight(185), document.PageSize.GetBottom(30));

            // Agregamos título de al medio del footer
            cb.BeginText();
            cb.SetFontAndSize(bf, 8);
            cb.SetTextMatrix(document.PageSize.GetRight(365), document.PageSize.GetBottom(40));
            cb.ShowText("Programa Operativo Anual " + PrintTime.Year);
            cb.EndText();
            cb.AddTemplate(footerTemplate, document.PageSize.GetRight(185), document.PageSize.GetBottom(30));

            // Agregamos númeracion de página
            cb.BeginText();
            cb.SetFontAndSize(bf, 8);
            cb.SetTextMatrix(document.PageSize.GetRight(53), document.PageSize.GetBottom(40));
            cb.ShowText(text);
            cb.EndText();
            float len = bf.GetWidthPoint(text, 8);
            cb.AddTemplate(footerTemplate, document.PageSize.GetRight(180) + len, document.PageSize.GetBottom(30));
            

            //Move the pointer and draw line to separate footer section from rest of page
            cb.MoveTo(25, document.PageSize.GetBottom(50));
            cb.LineTo(document.PageSize.Width - 25, document.PageSize.GetBottom(50));
            cb.Stroke();
            */
        }

        public override void OnCloseDocument(PdfWriter writer, Document document)
        {
            base.OnCloseDocument(writer, document);
            /*
            footerTemplate.BeginText();
            footerTemplate.SetFontAndSize(bf, 8);
            footerTemplate.SetTextMatrix(0, 0);
            footerTemplate.EndText();
            */
        }
    }
}
