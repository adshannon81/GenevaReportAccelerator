using System;
using System.Collections.Generic;
using Advent.Geneva.WFM.Framework.BaseImplementation;
using Advent.Geneva.WFM.Framework.Interfaces;
using Advent.Geneva.WFM.SQLDataAccess;
using Advent.Geneva.WFM.GenevaDataAccess;

namespace GenevaAutoReports
{
    /// <summary>
    /// Abstract class ActivityBase implements IActivity. Deriving from ActivityBase class instead of IActivity, gives a good starting point to implement an activity.
    /// Some of the methods that could be common across activity implementations are given a default definition. For overrides, the implemented methods are marked virtual. 
    /// Deriving from this class does not force you to implement the overloaded 'Run' method which takes a step name, in case you are not using activity steps.
    /// </summary>
    public class TemplateActivityBase : ActivityBase
    {
        /// <summary>
        /// <see cref="IActivity.Start">IActivity.Start</see>
        /// </summary>
        public override void Init()
        {
            base.Init(); //this will cache the activity settings.
            //you could have your code here to do other initializations.
        }

        /// <summary>
        /// <see cref="IActivity.Start">IActivity.Start</see>
        /// </summary>
        public override void Start()
        {
            //If it is a folder/sql server based activity, you could start the polling a folder, a sql server database table which could kick start activities.
            //Or it could be just initializing the DataTransform object.
        }

        /// <summary>
        /// <see cref="IActivity.Run(ActivityRun,IGenevaActions)">IActivity.Run(ActivityRun,IGenevaActions)</see>
        /// </summary>
        /// <param name="activityRun">ActivityRun instance</param>
        /// <param name="genevaInstance">GenevaAction instance</param>
        public override void Run(ActivityRun activityRun, IGenevaActions genevaInstance)
        {
            //This is where the actual activity run happens.
        }

        /// <summary>
        /// <see cref="IActivity.Stop">IActivity.Stop</see>
        /// </summary>
        public override void Stop()
        {
            //You could stop polling here.
            //You could dispose your system resources here or in ShutDown, depending on the resources and 
            //the way you have done w.r.t Init and Start above.
        }

        /// <summary>
        /// <see cref="IActivity.ShutDown">IActivity.ShutDown</see>
        /// </summary>
        public override void ShutDown()
        {
            //you could dispose of system resources here.
        }

        /// <summary>
        /// Validate if the Activity is ready to start.
        /// </summary>
        /// <returns>List of results</returns> 
        public override List<ValidateResult> Validate()
        {
            //Mainly make sure the settings has been correctly initialized, you have a http connection or sql connection etc,
            //so that the activity can be started.
            throw new NotImplementedException();
        }
    }
}
