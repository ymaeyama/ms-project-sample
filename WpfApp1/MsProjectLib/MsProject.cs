using Microsoft.Office.Interop.MSProject;
using System.Reflection;

namespace MsProjectLib
{
    public class MsProject
    {
        public void Load(string file)
        {
            object readOnly = false;
            var merge = PjMergeType.pjDoNotMerge;
            var pool = PjPoolOpen.pjDoNotOpenPool;
            object ignoreReadOnlyRecommended = false;

            var projectApp = new Application();
            projectApp.FileOpen(file, readOnly, merge, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                pool, Missing.Value, Missing.Value, ignoreReadOnlyRecommended, Missing.Value);
            Project proj = projectApp.ActiveProject;

            foreach (Task task in proj.Tasks)
            {
                if (task == null)
                {
                    continue;
                }

                var myTask = new MsProjectTask();
                myTask.Load(task);
            }

            projectApp.FileCloseAllEx(PjSaveType.pjSave);
            //projectApp.Quit(PjSaveType.pjSave);

            // 変更を保存しない場合
            //projectApp.FileCloseAllEx(PjSaveType.pjPromptSave);
        }
    }
}
