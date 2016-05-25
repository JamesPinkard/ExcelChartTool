﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class MeasurementRecordParser
    {
        public MeasurementRecordParser(MountainViewField previousField)
        {            
            _previousField = previousField;
        }

        public IEnumerable<MeasurementRecord> ProcessMeasurementRecord(QuarterTable quarter)
        {
            List<MeasurementRecord> records = new List<MeasurementRecord>();
            var uniqueWeekFieldQuery = new UniqueWeekFieldQuery();

            var quarterFields = quarter.GetFields();
            var weeks = uniqueWeekFieldQuery.GetUniqueWeekIndices(quarterFields);
            _previousField = quarterFields.First();



            return records;
        }

        public IEnumerable<MeasurementRecord> ProcessMeasurementRecord(WeekTable weekTable)
        {
            List<MeasurementRecord> records = new List<MeasurementRecord>();
            var weekFields = weekTable.GetFieldsForWeek();
            WeekMeasurementRecord weekMeasurementRecord = ProcessWeek(weekTable);
            records.Add(weekMeasurementRecord);
            

            if (weekFields.Count() > 1)
            {
                foreach (var measurement in weekFields.Skip(1))
                {
                    IndividualMeasurementRecord measurementRecord = ProcessField(measurement);
                    records.Add(measurementRecord);
                }
            }

            return records;
        }

        private IndividualMeasurementRecord ProcessField(MountainViewField measurement)
        {
            IndividualMeasurementRecord measurementRecord = ConvertToRecord(measurement);
            _previousField = measurement;
            return measurementRecord;
        }

        private WeekMeasurementRecord ProcessWeek(WeekTable weekTable)
        {
            var weekFields = weekTable.GetFieldsForWeek();
            var firstMeasurement = weekFields.First();
            IndividualMeasurementRecord firstMeasurementRecord = ProcessField(firstMeasurement);
            WeekMeasurementRecord weekMeasurementRecord = ConvertToWeeklyRecord(firstMeasurementRecord, weekTable);            
            return weekMeasurementRecord;
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
