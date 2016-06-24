using System;
using System.Collections.Generic;
using Advent.Geneva.WFM.Framework.BaseImplementation;
using Advent.Geneva.WFM.Framework.Interfaces;
using Advent.Geneva.WFM.GenevaDataAccess;
using Advent.Geneva.WFM.SQLDataAccess;

namespace GenevaAutoReports
{
    /// <summary>
    /// FolderPollingActivityBase class inherits from ActivityBase (see TemplateActivityBase) and implement some of the IActivity methods used commonly 
    /// by activities that polls a folder to kick off activity runs.
    /// </summary>
    class TemplateFolderPollingActivityBase : FolderPollingActivityBase
    {
        /// <summary>
        /// <see cref="IActivity.Start">IActivity.Start</see>
        /// </summary>
        public override void Start()
        {
            //start your polling here, see base.Start method.
            //base.Start(....);
            throw new NotImplementedException();
        }

        /// <summary>
        /// <see cref="IActivity.Run(ActivityRun,IGenevaActions)">IActivity.Run(ActivityRun,IGenevaActions)</see>
        /// </summary>
        /// <param name="activityRun">ActivityRun instance</param>
        /// <param name="genevaInstance">GenevaAction instance</param>
        public override void Run(ActivityRun activityRun, IGenevaActions genevaInstance)
        {
            //This will be invoked by the framework, this is where you implementation goes...where you may have to do a transformation and load
            //to Geneva.
            throw new NotImplementedException();
        }

        /// <summary>
        /// Validate if the Activity is ready to start.
        /// </summary>
        /// <returns>List of results</returns> 
        public override List<ValidateResult> Validate()
        {
            //Do all your validations here which should turn out with an empty list. see base static method ValidateFolderForReadAndWrite for help with folder validation.
            //ValidateFolderForReadAndWrite(...);

            //If there is no validation required leave the throw statement as it is. **Do not return a null list, return an empty list if there are no errors.
            throw new NotImplementedException();
        }
    }
}
