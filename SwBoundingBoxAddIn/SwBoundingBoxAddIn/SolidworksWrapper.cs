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
                        //need to add some logic for getting accurate points for asssembly
                        if (i == 0)
                            boundingBoxCoordinates = tempboxCoordinates;
                        

                        if (i > 0)
                        {
                           
                            boundingBoxCoordinates[0] = tempboxCoordinates[0] < boundingBoxCoordinates[0] ? tempboxCoordinates[0] : boundingBoxCoordinates[0];
                            
                            boundingBoxCoordinates[1] = tempboxCoordinates[1] < boundingBoxCoordinates[1] ? tempboxCoordinates[1] : boundingBoxCoordinates[1];
                           
                            boundingBoxCoordinates[2] = tempboxCoordinates[2] < boundingBoxCoordinates[2] ? tempboxCoordinates[2] : boundingBoxCoordinates[2];
                            
                            boundingBoxCoordinates[3] = tempboxCoordinates[3] > boundingBoxCoordinates[3] ? tempboxCoordinates[3] : boundingBoxCoordinates[3];
                            
                            boundingBoxCoordinates[4] = tempboxCoordinates[4] > boundingBoxCoordinates[4] ? tempboxCoordinates[4] : boundingBoxCoordinates[4];
                            
                            boundingBoxCoordinates[5] = tempboxCoordinates[5] > boundingBoxCoordinates[5] ? tempboxCoordinates[5] : boundingBoxCoordinates[5];

                        }
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

        public  void GetExtremePoints(ref double[] Coordinates, object[] partBodies)
        {
            double minX = 0, minY = 0, minZ = 0;
            double maxX = 0, maxY = 0, maxZ = 0;
            for (int i = 0; i < partBodies.Length; i++)
            {
                Body2 swBody = (Body2)partBodies[i];
                double x, y, z;
                //getting extreme coordinate in +ve x axis
                swBody.GetExtremePoint(1, 0, 0, out x, out y, out z);
                if (i == 0 || maxX < x)
                    maxX = x;
                //getting extreme coordinate in -ve x axis
                swBody.GetExtremePoint(-1, 0, 0, out x, out y, out z);
                if (i == 0 || minX > x)
                    minX = x;
                //getting extreme coordinate in +ve Y axis
                swBody.GetExtremePoint(0, 1, 0, out x, out y, out z);
                if (i == 0 || maxY < y)
                    maxY = y;
                //getting extreme coordinate in -ve Y axis
                swBody.GetExtremePoint(0, -1, 0, out x, out y, out z);
                if (i == 0 || minY > y)
                    minY = y;

                //getting extreme coordinate in +ve Z axis
                swBody.GetExtremePoint(0, 0, 1, out x, out y, out z);
                if (i == 0 || maxZ < z)
                    maxZ = z;

                //getting extreme coordinate in -ve Z axis
                swBody.GetExtremePoint(0, 0, -1, out x, out y, out z);
                if (i == 0 || minZ > z)
                    minY = z;

            }

            Coordinates[0] = minX;
            Coordinates[1] = minY;
            Coordinates[2] = minZ;

            Coordinates[3] = maxX;
            Coordinates[4] = maxY;
            Coordinates[5] = maxZ;
        }
    }
}
