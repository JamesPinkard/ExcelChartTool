using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    class MeasurementRecordParserSwap
    {
        public MeasurementRecordParserSwap(MountainViewField previousField)
        {
            _previousField = previousField;
        }

        public IEnumerable<MeasurementRecord> ProcessMeasurementRecord(QuarterTable quarter)
        {
            List<MeasurementRecord> records = new List<MeasurementRecord>();
            var uniqueWeekFieldQuery = new UniqueWeekFieldQuery();

            var quarterFields = quarter.GetFields();
            var weeks = uniqueWeekFieldQuery.GetUniqueWeekIndices(quarterFields);            

            foreach (var week in weeks)
            {
                var weekTable = quarter.GetTableForWeek(week);
                var measurementRecords = ProcessMeasurementRecord(weekTable);
                records.AddRange(measurementRecords);
            }

            var firstRecord = records.First() as WeekMeasurementRecord;
            records[0] = new QuarterMeasurementRecord(firstRecord, quarter.GetAverageWeeklyFlowRate());

            return records;
        }

        public IEnumerable<MeasurementRecord> ProcessMeasurementRecord(WeekTable weekTable)
        {
            List<MeasurementRecord> records = new List<MeasurementRecord>();
            var weekFields = weekTable.GetFieldsForWeek();


            foreach (var measurement in weekFields)
            {
                IndividualMeasurementRecord measurementRecord = ConvertToRecord(measurement);
                records.Add(measurementRecord);
                _previousField = measurement;
            }
            
            var firstRecord = records.First() as IndividualMeasurementRecord;
            records[0] = ConvertToWeeklyRecord(firstRecord, weekTable);

            return records;
        }

        private WeekMeasurementRecord ConvertToWeeklyRecord(IndividualMeasurementRecord measurementRecord, WeekTable weekTable)
        {
            var weekIndex = weekTable.GetWeek();
            var cumulativeTime = weekTable.GetCumulativeTime();
            var cumulativeFlow = weekTable.GetNetVolume();
            var averageFlow = weekTable.GetAverageWeeklyFlowRate();
            var weekMeasurementRecord = new WeekMeasurementRecord(measurementRecord, weekIndex, cumulativeTime, cumulativeFlow, averageFlow);
            return weekMeasurementRecord;
        }

        private IndividualMeasurementRecord ConvertToRecord(MountainViewField field)
        {
            var stationName = field.StationName;
            var measureTime = field.MeasureTime;
            var totalizerReading = field.TotalizerReading;
            var cumulativeTime = measureTime - _previousField.MeasureTime;
            var cumulativeFlow = totalizerReading - _previousField.TotalizerReading;
            var measurementRecord = new IndividualMeasurementRecord(stationName, measureTime, totalizerReading, cumulativeTime.TotalMinutes, cumulativeFlow);
            return measurementRecord;
        }

        private MountainViewField _previousField;
    }
}
