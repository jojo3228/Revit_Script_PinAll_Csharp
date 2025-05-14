using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;

namespace PinAll
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class AttachElements : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // Списки для хранения прикрепляемых объектов
            List<Element> levels = new List<Element>();
            List<Element> grids = new List<Element>();
            List<Element> links = new List<Element>();
            //BasePoint projectBasePoint = null;
            //BasePoint surveyPoint = null;

            // Счетчики прикрепленных объектов
            int attachedLevels = 0;
            int attachedGrids = 0;
            int attachedLinks = 0;
            int attachedProjectBasePoint = 0;
            int attachedSurveyPoint = 0;

            // Получаем уровни
            FilteredElementCollector levelCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(Level));
            levels.AddRange(levelCollector);

            // Получаем оси
            FilteredElementCollector gridCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(Grid));
            grids.AddRange(gridCollector);

            // Получаем связи
            FilteredElementCollector linkCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance));
            links.AddRange(linkCollector);

            // Получаем базовую точку проекта (Project Base Point)
            FilteredElementCollector projectBasePointCollector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_ProjectBasePoint)
                .OfClass(typeof(BasePoint));

            BasePoint projectBasePoint = projectBasePointCollector.FirstElement() as BasePoint;

            // Получаем точку съемки (Survey Point)
            FilteredElementCollector surveyPointCollector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_SharedBasePoint) // <-- используем другую категорию
                .OfClass(typeof(BasePoint));

            BasePoint surveyPoint = surveyPointCollector.FirstElement() as BasePoint;

            // Начинаем транзакцию
            using (Transaction transaction = new Transaction(doc, "Attach Levels, Grids, and Links"))
            {
                transaction.Start();

                // Проверяем и прикрепляем уровни
                foreach (Level level in levels)
                {
                    if (!IsPinned(level))
                    {
                        level.Pinned = true;
                        attachedLevels++;
                    }
                }

                // Проверяем и прикрепляем оси
                foreach (Grid grid in grids)
                {
                    if (!IsPinned(grid))
                    {
                        grid.Pinned = true;
                        attachedGrids++;
                    }
                }

                // Проверяем и прикрепляем связи
                foreach (RevitLinkInstance link in links)
                {
                    if (!IsPinned(link))
                    {
                        link.Pinned = true;
                        attachedLinks++;
                    }
                }

                // Проверяем и прикрепляем базовую точку проекта
                if (projectBasePoint != null && !IsPinned(projectBasePoint))
                {
                    projectBasePoint.Pinned = true;
                    attachedProjectBasePoint++;
                }

                // Проверяем и прикрепляем точку съемки
                if (surveyPoint != null && !IsPinned(surveyPoint))
                {
                    surveyPoint.Pinned = true;
                    attachedSurveyPoint++;
                }

                transaction.Commit();
            }

            // Формируем сообщение для пользователя
            List<string> messageParts = new List<string>();

            if (attachedLevels > 0) messageParts.Add($"- Уровни: {attachedLevels}");
            if (attachedGrids > 0) messageParts.Add($"- Оси: {attachedGrids}");
            if (attachedLinks > 0) messageParts.Add($"- Связи: {attachedLinks}");
            if (attachedProjectBasePoint > 0) messageParts.Add($"- Базовая точка проекта: {attachedProjectBasePoint}");
            if (attachedSurveyPoint > 0) messageParts.Add($"- Точка съемки: {attachedSurveyPoint}");

            string resultMessage = messageParts.Count > 0
                ? "Прикреплено:\n" + string.Join("\n", messageParts)
                : "Все уже было прикреплено.";

            TaskDialog.Show("Результат", resultMessage);
            return Result.Succeeded;
        }

        // Метод для проверки, прикреплен ли объект
        private bool IsPinned(Element element)
        {
            return element.Pinned;
        }
    }

}
