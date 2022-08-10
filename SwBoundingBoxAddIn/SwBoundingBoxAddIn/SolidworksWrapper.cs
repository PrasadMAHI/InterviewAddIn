using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SwBoundingBoxAddIn
{
    class SolidworksWrapper
    {
        private SldWorks SolidworksApplication { get; set; }

        public ModelDoc2 SolidworksModel { get; set; }
        
        /// <summary>
        /// Initializing the SolidWorksAplication of the wrapper
        /// </summary>
        /// <param name="swApp"></param>
        public SolidworksWrapper(SldWorks swApp)
        {
            SolidworksApplication = swApp;
        }
        /// <summary>
        /// Gets Active model(Part/assembly) from solidworks Application
        /// </summary>
        /// <returns></returns>
        public ModelDoc2 GetActiveModel()
        {
            //checking if there is no solidworks Application
            if (SolidworksApplication == null)
                return null;
            //Getting Active Model from the Application
            SolidworksModel = (ModelDoc2)SolidworksApplication.ActiveDoc;
            return SolidworksModel;
        }
        /// <summary>
        /// computes Bounding Box for Solidworks Document(Part/Assembly)
        /// </summary>

        public double[] ComputeBoundingBox()
        {
            if (SolidworksModel == null )
                return null;
          
            double[] boundingBoxCoordinates = new double[6];
            double[] tempboxCoordinates = new double[6];


            //checking if solidworks document is of Part Type
            if (SolidworksModel is PartDoc)
            {
                Object[] partBodies = null;
                PartDoc swPart = (PartDoc)SolidworksModel;
                //getting all bodies of an part 
                partBodies = (object[])swPart.GetBodies2((int)swBodyType_e.swSolidBody, true);

                //caluclates extreme points of the part from all bodies present in it 
                if (partBodies != null)
                    GetExtremePoints(ref boundingBoxCoordinates, partBodies);

            }
            //checking if solidworks document is of Assembly Type
            if (SolidworksModel is AssemblyDoc)
            {
                object[] swComponents= null;
                AssemblyDoc swAssembly = (AssemblyDoc)SolidworksModel;
                //getting components of the assembly
                swComponents = (object[])swAssembly.GetComponents(true);
                if(swComponents != null)
                {
                    for(int i = 0; i< swComponents.Length; i++)
                    {
                        Component2 component = null;
                        object[] componentBodies = null;
                        component = (Component2)swComponents[i];
                        if (component is AssemblyDoc)
                            continue;
                        //gets bodies present in a component
                        componentBodies = (object[])component.GetBodies2((int)swBodyType_e.swSolidBody);
                        if(componentBodies!=null)
                            GetExtremePoints(ref tempboxCoordinates, componentBodies);
                      
                    }
                }
            }
            return boundingBoxCoordinates;

        }
        /// <summary>
        /// gets extreme points for an particular part or component
        /// </summary>
        /// <param name="Coordinates"></param>
        /// <param name="partBodies"></param>

        public  void GetExtremePoints(ref double[] Coordinates, object[] partBodies,bool Iscomponent = false)
        {
            double minX = 0, minY = 0, minZ = 0;
            double maxX = 0, maxY = 0, maxZ = 0;
           
        }
    }
}
