using Microsoft.Office.Interop.MSProject;
using System;
using System.Collections.Generic;
using System.Linq;

namespace MsProjectLib
{
    public class MsProjectTask
    {
        public int Id { get; internal set; }
        public string Name { get; internal set; }
        public DateTime Start { get; internal set; }
        public DateTime Finish { get; internal set; }
        public string OutLineNumber { get; internal set; }
        public short OutlineLevel { get; internal set; }
        public int Index { get; internal set; }
        public string ParentOutLine { get; internal set; }
        public short PercentWorkComplete { get; internal set; }
        public List<MsProjectResource> Resoueces { get; internal set; }
        public bool Summary { get; internal set; }
        public string SummaryTaskName { get; internal set; }
        public List<MsProjectTask> Children { get; internal set; }

        public void Load(Task task)
        {
            this.Id = task.UniqueID;
            this.Name = task.Name;
            this.Start = task.Start;
            this.Finish = task.Finish;
            this.OutLineNumber = task.OutlineNumber;
            this.OutlineLevel = task.OutlineLevel;
            this.Index = task.Index;
            this.PercentWorkComplete = task.PercentWorkComplete;
            // = task.
            //this.Work = task.Work;

            if (task.OutlineLevel != 1)
            {
                this.ParentOutLine = task.OutlineParent.WBS.ToString();
                //ParentUid = data.GetParentUid(ParentOutLine, projectUid);
            }

            this.Summary = task.Summary;
            this.SummaryTaskName = ((dynamic)task).SummaryTaskName;

            this.Resoueces = new List<MsProjectResource>();
            foreach (Resource resource in task.Resources)
            {
                var myresource = new MsProjectResource();
                myresource.Id = resource.UniqueID;
                myresource.Name = resource.Name;
                this.Resoueces.Add(myresource);
            }

            this.Children = new List<MsProjectTask>();
            foreach (Task childTask in task.OutlineChildren)
            {
                var myChildTask = new MsProjectTask();
                myChildTask.Load(childTask);
                this.Children.Add(myChildTask);
            }
        }

        public override string ToString()
        {
            var resouceName = string.Join(", ", this.Resoueces.Select(x => x.Name));
            if (!string.IsNullOrEmpty(resouceName))
            {
                resouceName = $"[{resouceName}]";
            }
            return $"{this.OutLineNumber}: {this.Name}, Children = {this.Children.Count}{resouceName}";
        }
    }
}
