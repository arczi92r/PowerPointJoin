using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using System.IO;
using System.Linq;

namespace PowerPointJoin
{
    public class Program
    {
        public static int _uniqueId;
        static void Main(string[] args)
        {
            int copiedSlidePosition = 1;
            using (PresentationDocument destDoc = PresentationDocument.Open(@"C:\a.pptx", true))
            {

                using (var sourceDoc = PresentationDocument.Open(@"C:\b.pptx", false))
                {
                    
                    var destPresentationPart = destDoc.PresentationPart;
                    var destPresentation = destPresentationPart.Presentation;

                    var sourcePresentationPart = sourceDoc.PresentationPart;
                    var sourcePresentation = sourcePresentationPart.Presentation;
                    int countSlidesInSourcePresentation = sourcePresentation.SlideIdList.Count();

                    int copiedSlideIndex = (int)--copiedSlidePosition;

                    if (copiedSlideIndex < 0 || copiedSlideIndex >= countSlidesInSourcePresentation)
                        throw new ArgumentOutOfRangeException(nameof(copiedSlidePosition));

                    SlideId copiedSlideId = sourcePresentationPart.Presentation.SlideIdList.ChildElements[copiedSlideIndex] as SlideId;
                    SlidePart copiedSlidePart = sourcePresentationPart.GetPartById(copiedSlideId.RelationshipId) as SlidePart;

                    SlidePart addedSlidePart = destPresentationPart.AddPart<SlidePart>(copiedSlidePart);

                    SlideMasterPart addedSlideMasterPart = destPresentationPart.AddPart(addedSlidePart.SlideLayoutPart.SlideMasterPart);

                    UInt32 maxId = 1;
                    SlideId slideId = new SlideId
                    {
                        Id = maxId,
                        RelationshipId = destDoc.PresentationPart.GetIdOfPart(addedSlidePart)
                    };
                    destPresentation.SlideIdList.Append(slideId);

                    // Create new master slide ID
                    UInt32 _uniqueId = 1;
              
                    SlideMasterId slideMaterId = new SlideMasterId
                    {
                        Id = _uniqueId,
                        RelationshipId = destDoc.PresentationPart.GetIdOfPart(addedSlideMasterPart)
                    };
                    destDoc.PresentationPart.Presentation.SlideMasterIdList.Append(slideMaterId);

                    // change slide layout ID
                    FixSlideLayoutIds(destDoc.PresentationPart);

                    destDoc.PresentationPart.Presentation.Save();
                }
            

            }
        }
        public static void Copy(Stream sourcePresentationStream, uint copiedSlidePosition, Stream destPresentationStream)
        {
            using (var destDoc = PresentationDocument.Open(destPresentationStream, true))
            {
                var sourceDoc = PresentationDocument.Open(sourcePresentationStream, false);
                var destPresentationPart = destDoc.PresentationPart;
                var destPresentation = destPresentationPart.Presentation;

                //_uniqueId = GetMaxIdFromChild(destPresentation.SlideMasterIdList);
                //uint maxId = GetMaxIdFromChild(destPresentation.SlideIdList);

                var sourcePresentationPart = sourceDoc.PresentationPart;
                var sourcePresentation = sourcePresentationPart.Presentation;

                int copiedSlideIndex = (int)--copiedSlidePosition;

                int countSlidesInSourcePresentation = sourcePresentation.SlideIdList.Count();
                if (copiedSlideIndex < 0 || copiedSlideIndex >= countSlidesInSourcePresentation)
                    throw new ArgumentOutOfRangeException(nameof(copiedSlidePosition));

                SlideId copiedSlideId = sourcePresentationPart.Presentation.SlideIdList.ChildElements[copiedSlideIndex] as SlideId;
                SlidePart copiedSlidePart = sourcePresentationPart.GetPartById(copiedSlideId.RelationshipId) as SlidePart;

                SlidePart addedSlidePart = destPresentationPart.AddPart<SlidePart>(copiedSlidePart);

                SlideMasterPart addedSlideMasterPart = destPresentationPart.AddPart(addedSlidePart.SlideLayoutPart.SlideMasterPart);


                //// Create new slide ID
                //maxId++;
                //SlideId slideId = new SlideId
                //{
                //    Id = maxId,
                //    RelationshipId = destDoc.PresentationPart.GetIdOfPart(addedSlidePart)
                //};
                //destPresentation.SlideIdList.Append(slideId);

                //// Create new master slide ID
                //_uniqueId++;
                //SlideMasterId slideMaterId = new SlideMasterId
                //{
                //    Id = _uniqueId,
                //    RelationshipId = destDoc.PresentationPart.GetIdOfPart(addedSlideMasterPart)
                //};
                //destDoc.PresentationPart.Presentation.SlideMasterIdList.Append(slideMaterId);

                // change slide layout ID
                FixSlideLayoutIds(destDoc.PresentationPart);

                destDoc.PresentationPart.Presentation.Save();
            }
            sourcePresentationStream.Close();
            destPresentationStream.Close();
        }
        public static void FixSlideLayoutIds(PresentationPart presPart)
        {
            // Make sure that all slide layouts have unique ids.
            foreach (SlideMasterPart slideMasterPart in presPart.SlideMasterParts)
            {
                UInt32 uniqueId = 1;
                foreach (SlideLayoutId slideLayoutId in slideMasterPart.SlideMaster.SlideLayoutIdList)
                {
                   
                    slideLayoutId.Id = (uint)uniqueId;
                    uniqueId++;
                }

                slideMasterPart.SlideMaster.Save();
            }
        }
        private static string GetSlideLayoutType(SlideLayoutPart slideLayoutPart)
        {
            CommonSlideData slideData = slideLayoutPart.SlideLayout.CommonSlideData;

            return slideData.Name;
        }







        // Insert the specified slide into the presentation at the specified position.
        //public static void InsertNewSlide(PresentationDocument presentationDocument, int position, string slideTitle)
        //{
        //    if (presentationDocument == null)
        //    {
        //        throw new ArgumentNullException("presentationDocument");
        //    }

        //    if (slideTitle == null)
        //    {
        //        throw new ArgumentNullException("slideTitle");
        //    }

        //    PresentationPart presentationPart = presentationDocument.PresentationPart;

        //    // Verify that the presentation is not empty.
        //    if (presentationPart == null)
        //    {
        //        throw new InvalidOperationException("The presentation document is empty.");
        //    }

        //    // Declare and instantiate a new slide.
        //    Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
        //    uint drawingObjectId = 1;

        //    // Construct the slide content.            
        //    // Specify the non-visual properties of the new slide.
        //    D.NonVisualGroupShapeProperties nonVisualProperties = slide.CommonSlideData.ShapeTree.AppendChild(new D.NonVisualGroupShapeProperties());
        //    nonVisualProperties.NonVisualDrawingProperties = new D.NonVisualDrawingProperties() { Id = 1, Name = "" };
        //    nonVisualProperties.NonVisualGroupShapeDrawingProperties = new D.NonVisualGroupShapeDrawingProperties();
        //    nonVisualProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

        //    // Specify the group shape properties of the new slide.
        //    slide.CommonSlideData.ShapeTree.AppendChild(new GroupShapeProperties());

        //}
        //    public static void CreatePresentation(string filepath)
        //{
        //    // Create a presentation at a specified file path. The presentation document type is pptx, by default.
        //    PresentationDocument presentationDoc = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation);
        //    PresentationPart presentationPart = presentationDoc.AddPresentationPart();
        //    presentationPart.Presentation = new Presentation();

        //    CreatePresentationParts(presentationPart);

        //    // Close the presentation handle
        //    presentationDoc.Close();
        //}


        //private static void CreatePresentationParts(PresentationPart presentationPart)
        //{
        //    SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" });
        //    SlideIdList slideIdList1 = new SlideIdList(new SlideId() { Id = (UInt32Value)256U, RelationshipId = "rId2" });
        //    SlideSize slideSize1 = new SlideSize() { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 };
        //    NotesSize notesSize1 = new NotesSize() { Cx = 6858000, Cy = 9144000 };
        //    DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

        //    presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1);

        //    SlidePart slidePart1;
        //    SlideLayoutPart slideLayoutPart1;
        //    SlideMasterPart slideMasterPart1;
        //    ThemePart themePart1;


        //    slidePart1 = CreateSlidePart(presentationPart);
        //    slideLayoutPart1 = CreateSlideLayoutPart(slidePart1);
        //    slideMasterPart1 = CreateSlideMasterPart(slideLayoutPart1);
        //    themePart1 = CreateTheme(slideMasterPart1);

        //    slideMasterPart1.AddPart(slideLayoutPart1, "rId1");
        //    presentationPart.AddPart(slideMasterPart1, "rId1");
        //    presentationPart.AddPart(themePart1, "rId5");
        //}
    }
}
