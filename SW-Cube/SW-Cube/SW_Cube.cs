using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using System.IO;
using System.IO.Ports;
using System.Windows.Forms;
using System.Windows.Media.Media3D;
using System.Runtime.InteropServices;
using System.Threading;
using System.Diagnostics;

using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swcommands;
//using SolidWorksTools;

namespace SW_Cube
{

    public class SW_Cube : ISwAddin
    {
        public SldWorks _swApp;
        MathUtility swMathUtility;
        MathTransform orientation;
        private int mSWCookie;
        int solidFrameInterval;
        System.Threading.Timer solidFrameTimerT;

        //solid vars
        double MAX_THETA_DIFF_UNLOCK = 0.01;
        double MAX_AXIS_DIFF_UNLOCK = 0.0001;

        int sFPS = 60;
        bool _solidStartedByForm = false;
        bool solidRunning = false;
        bool solidDoc = false;
        bool solidFrameTimerTEnabled = false;
        static Mutex solidFrameMutex = new Mutex();
        bool solidMovement = false;
        bool mpuStable = true;
        Quaternion lastLockedQuat = new Quaternion(0, 0, 0, 1);
        Quaternion quat = new Quaternion(0, 0, 0, 1);
        Quaternion invCalQuat = new Quaternion(0, 0, 0, 1); //Vect= (0, 1, 0),  Angle=0. Identity.
        Quaternion worldQuat = new Quaternion(0, 0, 0, 1); //Vect= (0, 1, 0),  Angle=0. Identity.
                                                           //receive array from cube
        bool calAlgNum2 = false;

        [ComRegisterFunction()]
        private static void ComRegister(Type t)
        {
            string keyPath = String.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);
            var baseReg = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry64);
            //32 registry. goes to HKEY_LOCALMACHINE/SOFTWARE/WOW6432Node on 64 bit machines.
            //using (Microsoft.Win32.RegistryKey rk = Microsoft.Win32.Registry.CreateSubKey(keyPath))
            using (Microsoft.Win32.RegistryKey rk = baseReg.CreateSubKey(keyPath))
            {
                rk.SetValue(null, 1); // Load at startup
                rk.SetValue("Title", "My SwAddin"); // Title
                rk.SetValue("Description", "All your pixels belong to us2"); // Description
            }
        }

        [ComUnregisterFunction()]
        private static void ComUnregister(Type t)
        {
            string keyPath = String.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);
            var baseReg = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry64);
            baseReg.DeleteSubKeyTree(keyPath);
        }

        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            _swApp = (SldWorks)ThisSW;
            mSWCookie = Cookie;
            // Set-up add-in call back info
            bool result = _swApp.SetAddinCallbackInfo(0, this, Cookie);
            swMathUtility = (MathUtility)_swApp.GetMathUtility();
            orientation = swMathUtility.CreateTransform(new double[1]);
            this.UISetup();
            _swApp.SendMsgToUser2("timer started",
(int)swMessageBoxIcon_e.swMbInformation,
(int)swMessageBoxBtn_e.swMbOk);
            return true;
        }

        public bool DisconnectFromSW()
        {
            this.UITeardown();
            return true;
        }

        private void UISetup()
        {

            solidFrameInterval = (int)(1000 / sFPS);
            solidFrameTimerT = new System.Threading.Timer(SolidFrameT, null, solidFrameInterval, solidFrameInterval);
        }

        private void UITeardown()
        {
        }

        //method for updating the solid cam view
        private void SolidFrameT(object myObject)//Vector3D a, Double theta)
        {
            Stopwatch stopWatch = new Stopwatch();
            double[] times = new double[8];
            stopWatch.Start();

            if (solidFrameMutex.WaitOne(0))
            {
                //no update over "noise", no update during calibration
                solidMovement = MovementFilter();
                if (!solidMovement || !mpuStable)
                {
                    solidFrameMutex.ReleaseMutex();
                    return;
                }

                lastLockedQuat = quat;
                Quaternion tempQuat = GetCorrectedQuat();
                Vector3D a = tempQuat.Axis;
                double theta = tempQuat.Angle;
                theta *= Math.PI / 180;
                //move object instead of the camera
                theta = -theta;

                //0 ms
                times[0] = stopWatch.ElapsedMilliseconds;
                try
                {
                    if (!solidDoc)
                    {
                        //5-19 ms
                        if (_swApp.ActiveDoc != null)
                        {
                            solidDoc = true;
                        }
                    }
                    //avoiding exceptions if possible                        
                    if (solidDoc)
                    {
                        times[1] = stopWatch.ElapsedMilliseconds;
                        //5-14 ms
                        IModelDoc doc = _swApp.ActiveDoc;
                        try
                        {
                            times[2] = stopWatch.ElapsedMilliseconds;
                            //4-6 ms somehow solid won't allow this to happen at once
                            IModelView view = doc.ActiveView;
                            times[3] = stopWatch.ElapsedMilliseconds;
                            tempQuat.Invert();
                            double[,] rotation = QuatToRotation(tempQuat);
                            //TODO: translate :(
                            //15-23 ms no need to translate just yet!
                            //MathTransform translate = view.Translation3;
                            //TODO: rescale :(
                            //no need to rescale yet either
                            //double scale = view.Scale2;
                            times[4] = stopWatch.ElapsedMilliseconds;
                            double[] tempArr = new double[16];
                            //new X axis
                            tempArr[0] = rotation[0, 0];
                            tempArr[1] = rotation[1, 0];
                            tempArr[2] = rotation[2, 0];
                            //new Y axis
                            tempArr[3] = rotation[0, 1];
                            tempArr[4] = rotation[1, 1];
                            tempArr[5] = rotation[2, 1];
                            //new Z axis
                            tempArr[6] = rotation[0, 2];
                            tempArr[7] = rotation[1, 2];
                            tempArr[8] = rotation[2, 2];
                            //translation - doesn't mater for orientation!
                            tempArr[9] = 0;
                            tempArr[10] = 0;
                            tempArr[11] = 0;
                            //scale - doesn't mater for orientation!
                            tempArr[12] = 1;
                            //?
                            tempArr[13] = 0;
                            tempArr[14] = 0;
                            tempArr[15] = 0;
                            //? ms
                            orientation.ArrayData = tempArr;
                            times[5] = stopWatch.ElapsedMilliseconds;
                            //? ms
                            view.Orientation3 = orientation;
                            times[6] = stopWatch.ElapsedMilliseconds;
                            //? ms
                            view.RotateAboutCenter(1, 1);
                            //view.GraphicsRedraw(new int[] { });
                            times[7] = stopWatch.ElapsedMilliseconds;

                        }
                        //no active view
                        catch (Exception ex)
                        {
                            solidDoc = false;
                            //MessageBox.Show("Unable to rotate Solid Camera!\n" + ex.ToString());
                        }
                    }
                }
                //no _swApp
                catch (Exception ex)
                {
                    solidRunning = false;
                    MessageBox.Show("Oh no! Something went wrong with Solid!\n" + ex.ToString());
                }
                solidFrameMutex.ReleaseMutex();
                stopWatch.Stop();
            }
        }

        //method for getting the corrected current quat
        public Quaternion GetCorrectedQuat()
        {
            //tempQuat = R(C^-1)
            Quaternion tempQuat = Quaternion.Multiply(invCalQuat, quat);
            if (calAlgNum2)
            {
                //World view correction:
                //tempQuat = (C^-1)R
                tempQuat = Quaternion.Multiply(quat, invCalQuat);
                //tempQuat = W(C^-1)R
                tempQuat = Quaternion.Multiply(tempQuat, worldQuat);
                Quaternion invWorldQuat = new Quaternion(worldQuat.X, worldQuat.Y, worldQuat.Z, worldQuat.W);
                invWorldQuat.Invert();
                //tempQuat = W(C^-1)R(W^-1)
                tempQuat = Quaternion.Multiply(invWorldQuat, tempQuat);
            }
            return tempQuat;
        }

        //TODO: make a good filter.
        //function that checks whether a an actual movement of the cube was made
        private bool MovementFilter()
        {
            double diffTheta = lastLockedQuat.Angle - quat.Angle;
            Vector3D diffVector = Vector3D.Subtract(lastLockedQuat.Axis, quat.Axis);
            if (!(diffTheta > MAX_THETA_DIFF_UNLOCK || diffVector.Length > MAX_AXIS_DIFF_UNLOCK))
            {
                ////TODO: dis/re-enable timers?
                ////TODO: recalibrate to prevent stationary drift of cube over time?
                //avoid jumping due to drifting
                lastLockedQuat = quat;
                //inventorFrameTimer.Stop();
                return false;
            }
            return true;
        }

        public double[,] QuatToRotation(Quaternion a)
        {
            double[,] rotation = new double[3, 3];
            rotation[0, 0] = 1 - (2 * a.Y * a.Y + 2 * a.Z * a.Z);
            rotation[0, 1] = 2 * a.X * a.Y + 2 * a.Z * a.W;
            rotation[0, 2] = 2 * a.X * a.Z - 2 * a.Y * a.W;

            rotation[1, 0] = 2 * a.X * a.Y - 2 * a.Z * a.W;
            rotation[1, 1] = 1 - (2 * a.X * a.X + 2 * a.Z * a.Z);
            rotation[1, 2] = 2 * a.Y * a.Z + 2 * a.X * a.W;

            rotation[2, 0] = 2 * a.X * a.Z + 2 * a.Y * a.W;
            rotation[2, 1] = 2 * a.Y * a.Z - 2 * a.X * a.W;
            rotation[2, 2] = 1 - (2 * a.X * a.X + 2 * a.Y * a.Y);

            return rotation;
        }

    }
}
