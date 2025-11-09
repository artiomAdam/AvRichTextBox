using Avalonia.Controls;
using Avalonia.Input;
using DocumentFormat.OpenXml.VariantTypes;
using RtfDomParserAv;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using static AvRichTextBox.FlowDocument;

namespace AvRichTextBox;

public partial class RichTextBox
{
    
   private void ToggleItalics()
   {
      if (IsReadOnly) return;
      FlowDoc.ToggleItalic();

   }

   private void ToggleBold()
   {
      if (IsReadOnly) return;
      FlowDoc.ToggleBold();

   }

   private void ToggleUnderlining()
   {
      if (IsReadOnly) return;
      FlowDoc.ToggleUnderlining();

   }

   private void CopyToClipboard()
   {      
      
      var dataObject = new DataObject();

      //create rtf string
      List<IEditable> newInlines = FlowDoc.GetRangeInlines(FlowDoc.Selection);
      string rtfString = RtfConversions.GetRtfFromInlines(newInlines);
      byte[] rtfbytes = System.Text.Encoding.Default.GetBytes(rtfString);

      dataObject.Set("Rich Text Format", rtfbytes);
      dataObject.Set("Text", FlowDoc.Selection.GetText());
            
      TopLevel.GetTopLevel(this)!.Clipboard!.SetDataObjectAsync(dataObject);
      
   }

   
   private async void PasteFromClipboard()
   {
      if (IsReadOnly) return;

      bool TextPasted = false;
      int originalSelectionStart = FlowDoc.Selection.Start;
      int newSelPoint = originalSelectionStart;

      string[] formats = await TopLevel.GetTopLevel(this)!.Clipboard!.GetFormatsAsync();
      if (formats.Contains ("Rich Text Format"))
      {
         object? rtfobj = await TopLevel.GetTopLevel(this)!.Clipboard!.GetDataAsync("Rich Text Format");
         if (rtfobj != null)
         {
            byte[] rtfbytes = (byte[])rtfobj;
            string rtfstring = System.Text.Encoding.Default.GetString(rtfbytes!);

            RTFDomDocument dom = new();
            dom.LoadRTFText(rtfstring);
            List<IEditable> insertInlines = RtfConversions.GetInlinesFromRtf(dom);
            insertInlines.Reverse();
            int addedchars = FlowDoc.SetRangeToInlines(FlowDoc.Selection, insertInlines);

            newSelPoint = Math.Min(newSelPoint + addedchars, FlowDoc.DocEndPoint - 1);

            TextPasted = true;
         }
      }
      else if (formats.Contains("Text"))
      {
         object? textobj = await TopLevel.GetTopLevel(this)!.Clipboard!.GetDataAsync("Text");

         if (textobj != null)
         {
            string pasteText = textobj.ToString()!;
            FlowDoc.SetRangeToText(FlowDoc.Selection, pasteText);

            newSelPoint = Math.Min(newSelPoint + pasteText.Length, FlowDoc.DocEndPoint - 1);

            TextPasted = true;
         }
      }

      /*//Change by: Arty Adam
      else if(formats.Contains("PNG"))
      {
         object? pngObj = await TopLevel.GetTopLevel(this)!.Clipboard!.GetDataAsync("PNG");

         if (pngObj is byte[] pngBytes && pngBytes.Length > 0)
         {
            // Convert PNG bytes to Avalonia Bitmap
            using var ms = new MemoryStream(pngBytes);
            var bitmap = new Avalonia.Media.Imaging.Bitmap(ms);

            // Create an inline image element (depends on your FlowDoc model)
            var imageInline = new InlineImage(bitmap)
            {
               Width = bitmap.PixelSize.Width,
               Height = bitmap.PixelSize.Height
            };

            // Insert the image into the flow document
            FlowDoc.SetRangeToInlines(FlowDoc.Selection, new List<IEditable> { imageInline });

            newSelPoint = Math.Min(newSelPoint + 1, FlowDoc.DocEndPoint - 1);
            TextPasted = true;
         }*/

      if (TextPasted)
      {
         this.DocIC.UpdateLayout();
         await Task.Delay(100); //necessary for following operations

         FlowDoc.Selection.EndParagraph.CallRequestInlinesUpdate();  // important
         FlowDoc.Selection.EndParagraph.UpdateEditableRunPositions();

         FlowDoc.Select(newSelPoint, 0);
         FlowDoc.UpdateSelection();

         FlowDoc.Selection.BiasForwardStart = false;
         FlowDoc.Selection.BiasForwardEnd = false;
         FlowDoc.SelectionExtendMode = ExtendMode.ExtendModeNone;

         CreateClient();


      }

   }


}
