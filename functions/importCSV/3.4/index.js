import { now } from 'lodash';
import Papa from 'papaparse';
import { read, utils } from 'xlsx/xlsx.mjs';
import { createOne, updateOne, deleteOne } from '../../utils/mutations';
import { getOne, getAll } from '../../utils/queries';
import {
  snakeToCamel,
  convertToDBDateFormat,
  convertToBoolean,
  throwError,
} from '../../utils/helpers';

const splitArray = (arr, size) => {
  if (size <= 0) {
    return arr;
  }

  const result = [];
  for (let i = 0; i < arr.length; i += size) {
    result.push(arr.slice(i, i + size));
  }

  return result;
};

const getImportLines = async (fileUrl, fileType, csvDelimiter, logging) => {
  let data = [];
  const startTime = now();

  if (fileType === 'csv') {
    const csvText = await (await fetch(fileUrl)).text();
    const { data: rows, errors } = Papa.parse(csvText, {
      header: true,
      skipEmptyLines: true,
      delimiter:
        csvDelimiter && csvDelimiter !== 'auto' ? csvDelimiter : '',
      transformHeader: (h) => h.trim(),
    });
    if (logging && errors && errors.length) {
      console.warn(
        `CSV parse warnings: ${JSON.stringify(errors.slice(0, 5))}`,
      );
    }
    data = rows;
  } else {
    const { buffer } = await (await fetch(fileUrl)).blob();
    const workbook = read(buffer, { raw: true, dense: true, cellDates: true });
    data = utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
  }

  if (logging) {
    console.log(`Retrieved the import file in: ${now() - startTime} ms`);
  }

  return data;
};

const formatImportLineValues = (
  importLinesToSanitize,
  propertyMappings,
  propertyMappingsFormat,
) => {
  return importLinesToSanitize.map((importLine) => {
    // convert import values to database acceptable formats for text, decimal, number, date, date/time, time, decimal and checkbox properties
    propertyMappings.forEach((mapping) => {
      const importLineValue = mapping.key.endsWith('*')
        ? importLine[mapping.key] || importLine[mapping.key.slice(0, -1)] // Remove the last character (the asterisk)
        : importLine[mapping.key];
      propertyMappingsFormat.forEach((formatMapping) => {
        if (mapping.key === formatMapping.key && importLineValue) {
          const formatArray = formatMapping.value.trim().split(',');
          switch (formatArray[0].toString().toLowerCase()) {
            case 'text':
              importLine[mapping.key] = `${importLineValue}`;
              break;
            case 'decimal':
            case 'price':
              // if there is no dot notation, fix this by adding the .00
              const sanitizedDecimalValue = importLineValue
                .toString()
                .replace(',', '.');
              if (!isNaN(parseFloat(sanitizedDecimalValue))) {
                if (sanitizedDecimalValue.indexOf('.') === -1)
                  importLine[mapping.key] = sanitizedDecimalValue + '.00';
                else {
                  importLine[mapping.key] = parseFloat(sanitizedDecimalValue)
                    .toFixed(2)
                    .toString();
                }
              } else {
                importLine[mapping.key] = '';
              }
              break;
            case 'number':
              if (isNaN(parseFloat(importLineValue))) {
                importLine[mapping.key] = '';
              }
              break;
            case 'checkbox':
              importLine[mapping.key] = convertToBoolean(importLineValue);
              break;
            default: // date/time formats
              if (formatArray)
                importLine[mapping.key] = convertToDBDateFormat(
                  importLineValue,
                  formatArray[1] ? formatArray[1] : 'dd-MM-yyyy',
                  formatArray[0] ? formatArray[0] : 'Date',
                );
              break;
          }
        }
      });
    });
    return importLine;
  });
};

const prepareImportLines = async (
  importLines,
  existingRecords,
  uniqueRecordColumnName,
  deduplicate,
  defaultMappings,
  propertyMappings,
  propertyMappingsUpdate,
  uniqueRecordIdentifier,
  relationLookupData,
) => {
  const recordsToCreate = [];
  const recordsToUpdate = [];

  importLines.forEach((importLine) => {
    const importObj = {};
    const updateObj = {};

    defaultMappings.forEach((mapping) => {
      const dbName = snakeToCamel(mapping.key);
      importObj[dbName] = mapping.value;
      updateObj[dbName] = mapping.value;
    });

    propertyMappings.forEach((mapping) => {
      const importLineValue = mapping.key.endsWith('*')
        ? importLine[mapping.key] || importLine[mapping.key.slice(0, -1)] // Remove the last character (the asterisk)
        : importLine[mapping.key];
      if (mapping.isRelation) {
        const relationalLookupData = relationLookupData.find(
          (item) => item.relationImportName === mapping.key,
        );
        if (relationalLookupData) {
          const { records } = relationalLookupData;
          const relationObject = records.find(
            (element) => element[mapping.value] == importLineValue,
          );

          if (relationObject) {
            importObj[mapping.relationName] = relationObject.id;
          } else {
            importObj[mapping.relationName] = {};
          }
        }
      } else {
        importObj[mapping.value] = importLineValue ?? null;
      }
    });

    propertyMappingsUpdate.forEach((mapping) => {
      const importLineValue = mapping.key.endsWith('*')
        ? importLine[mapping.key] || importLine[mapping.key.slice(0, -1)] // Remove the last character (the asterisk)
        : importLine[mapping.key];
      updateObj[mapping.value] = importLineValue ?? null;
    });

    if (deduplicate) {
      if (importLine[uniqueRecordColumnName] !== '') {
        const existingRecord = existingRecords.find(
          (record) =>
            record[uniqueRecordIdentifier.value].toString() ===
            importLine[uniqueRecordIdentifier.key].toString(),
        );
        if (existingRecord) {
          if (propertyMappingsUpdate.length > 0) {
            updateObj.id = existingRecord.id;
            recordsToUpdate.push(updateObj);
          } else {
            importObj.id = existingRecord.id;
            recordsToUpdate.push(importObj);
          }
        } else {
          recordsToCreate.push(importObj);
        }
      }
    } else {
      recordsToCreate.push(importObj);
    }
  });
  return { recordsToUpdate, recordsToCreate };
};

const prepareRelationMappings = (propertyMapping, formatMappings) => {
  const enrichedMapping = propertyMapping.map((mapping) => {
    mapping.isRelation = false;
    if (mapping.value.indexOf('.') !== -1) {
      const relationModelProperty = mapping.value.split('.');
      if (relationModelProperty && relationModelProperty.length === 2) {
        const modelName = snakeToCamel(relationModelProperty[0]);
        const propertyName = snakeToCamel(relationModelProperty[1]);
        mapping.isRelation = true;
        mapping.relationModelName =
          modelName.charAt(0).toUpperCase() + modelName.slice(1);
        mapping.relationName = modelName;
        mapping.relationImportName = mapping.key;
        mapping.value = propertyName;
        const mappingPropertyFormat = formatMappings.find(
          (formatMap) =>
            mapping.key === formatMap.key ||
            mapping.key === formatMap.key.slice(0, -1), // Remove the last character (the asterisk)
        );
        if (mappingPropertyFormat) {
          mapping.relationPropertyType = mappingPropertyFormat.value
            .toLowerCase()
            .trim();
        }
        return mapping;
      }
    } else {
      mapping.value = snakeToCamel(mapping.value);
      return mapping;
    }
  });
  return enrichedMapping || propertyMapping;
};

const getWhere = (importLines, fileColumn, dbColumn, isDecimal) => {
  const uniqueValues = [];
  importLines.forEach((item) => {
    if (item[fileColumn]) {
      if (!uniqueValues.includes(item[fileColumn]))
        uniqueValues.push(item[fileColumn]);
    }
  });
  if (isDecimal) {
    return { _or: uniqueValues.map((item) => ({ [dbColumn]: { eq: item } })) };
  } else {
    return { [dbColumn]: { in: uniqueValues } };
  }
};

const getRelationLookupData = async (
  mainMappings,
  updateMappings,
  importLines,
) => {
  const mappings = mainMappings.concat(updateMappings);
  const lookupData = [];

  for (let i = 0; i < mappings.length; i++) {
    const mapping = mappings[i];
    if (mapping.isRelation) {
      const isMappingDecimal = mapping.relationPropertyType === 'decimal';
      const whereFilter = getWhere(
        importLines,
        mapping.key,
        mapping.value,
        isMappingDecimal,
      );

      const gqlRelationQuery = `{
        all${mapping.relationModelName}(skip: $skip, take: $take, where: $where) {
          results {
            id
            ${mapping.value}
          }
          totalCount
        }
      }`;

      const returnedData = await getAll(gqlRelationQuery, 0, 200, [], {
        where: whereFilter,
      });
      lookupData.push({
        relationImportName: mapping.key,
        relationDBName: mapping.relationImportName,
        records: returnedData,
      });
    }
  }
  return lookupData;
};

const processImportLines = async (
  importLines,
  propertyMappingMain,
  defaultMappings,
  propertyMappings,
  propertyMappingsFormat,
  propertyMappingsUpdateSpecific,
  deduplicate,
  uniqueRecordColumnName,
  uniqueRecordColumnType,
  batched,
  batchSize,
  batchOffset,
  startTime,
  logging,
  modelName,
) => {
  let allCurrentRecords = [];
  let loggingMessage = `Finished batch ${batchOffset + 1} (import lines: ${
    batchOffset * batchSize
  } to ${batchOffset * batchSize + importLines.length})`;

  const formattedImportLines = formatImportLineValues(
    importLines,
    propertyMappings,
    propertyMappingsFormat,
  );

  const relationLookupData = await getRelationLookupData(
    propertyMappingMain,
    propertyMappingsUpdateSpecific,
    formattedImportLines,
  );

  const relationLookupDataString = JSON.stringify(relationLookupData);

  if (logging && relationLookupDataString !== '[]') {
    console.log(
      'Relational lookup data (first 2000 characters): ' +
        relationLookupDataString.substring(0, 2000),
    );
  }

  const propertiesDBNames = [];
  propertyMappingMain.forEach((item) => {
    if (!item.isRelation) propertiesDBNames.push(item.value);
  });

  let uniqueRecordIdentifier = undefined;

  if (deduplicate && uniqueRecordColumnName !== '') {
    uniqueRecordIdentifier = propertyMappings.find(
      (item) =>
        item.key.toString().toLowerCase().trim() ===
        uniqueRecordColumnName.toString().toLowerCase().trim(),
    );

    if (uniqueRecordIdentifier === undefined)
      throwError(
        'No unique identifier can be found in the mappings. Make sure you add the import column and property database name (in snake_case) in the mappings above!',
      );
  }

  if (deduplicate && formattedImportLines && formattedImportLines.length > 0) {
    const whereFilter = getWhere(
      formattedImportLines,
      uniqueRecordColumnName,
      uniqueRecordIdentifier.value,
      uniqueRecordColumnType === 'decimal',
    );

    const gqlQuery = `{
      all${modelName}(skip: $skip, take: $take, where: $where) {
        results {
          id
          ${propertiesDBNames.join(' ')}
        }
        totalCount
      }
    }`;

    allCurrentRecords = await getAll(gqlQuery, 0, 200, [], {
      where: whereFilter,
    });
  }

  const processedImportLines = await prepareImportLines(
    formattedImportLines,
    allCurrentRecords,
    uniqueRecordColumnName,
    deduplicate,
    defaultMappings,
    propertyMappingMain,
    propertyMappingsUpdateSpecific,
    uniqueRecordIdentifier,
    relationLookupData,
  );
  const { recordsToCreate, recordsToUpdate } = processedImportLines;

  if (recordsToCreate && recordsToCreate.length > 0) {
    const createQuery = `
      mutation {
        createMany${modelName}(input: $input) {
          id
        }
      }
    `;

    if (logging && !batched) {
      console.log(
        `Collection to create: ${
          recordsToCreate.length
        } items (first 2000 characters): ${JSON.stringify(
          recordsToCreate,
        ).substring(0, 2000)}`,
      );
    }

    // The createMany mutation does not accept an ID property so we remove this from the records
    const sanitizedRecordsToCreate = recordsToCreate.map((record) => {
      delete record.id;
      return record;
    });

    const { data: createdData, errors: createdErrors } = await gql(
      createQuery,
      { input: sanitizedRecordsToCreate },
    );

    if (createdErrors) throwError(createdErrors);

    const createdIdsCol = Object.values(createdData)[0];
    if (!batched)
      console.log('Finished creating ' + createdIdsCol.length + ' records');
    else loggingMessage += ` New: ${createdIdsCol.length}`;
  }

  if (recordsToUpdate && recordsToUpdate.length > 0) {
    if (logging && !batched) {
      console.log(
        `Collection to update: ${
          recordsToUpdate.length
        } items (first 2000 characters): ${JSON.stringify(
          recordsToUpdate,
        ).substring(0, 2000)}`,
      );
    }
    const updateQuery = `
      mutation {
        upsertMany${modelName}(input: $input) {
          id
        }
      }
    `;

    const { data: updatedData, errors: updatedErrors } = await gql(
      updateQuery,
      { input: recordsToUpdate },
    );

    if (updatedErrors) throwError(updatedErrors);

    if (logging && updatedData) {
      const updatedIdsCol = Object.values(updatedData)[0];
      if (!batched)
        console.log('Finished updating ' + updatedIdsCol.length + ' records');
      else loggingMessage += ` Updated: ${updatedIdsCol.length}`;
    }
  }

  if (logging) {
    if (!batched) console.log(`Import finished in: ${now() - startTime} ms`);
    else loggingMessage += ` in: ${now() - startTime} ms`;
  }

  if (!batched)
    loggingMessage = `Records created: ${recordsToCreate.length}, records updated: ${recordsToUpdate.length}`;

  return loggingMessage;
};

const importFile = async ({
  fileUrl,
  fileType,
  csvDelimiter,
  model: { name: modelName },
  uniqueRecordColumnName,
  uniqueRecordColumnType,
  deduplicate,
  propertyMappings = [],
  propertyMappingsUpdate = [],
  propertyMappingsFormat = [],
  defaultMappings = [],
  batched,
  batchModel,
  batchSize,
  batchSizeProperty,
  batchOffsetProperty,
  batchFileNameProperty,
  logging,
  validateRequiredColumns,
}) => {
  try {
    const batchModelName = batchModel ? batchModel.name : null;
    const batchSizePropertyName = batchSizeProperty
      ? batchSizeProperty[0].name
      : null;
    const batchOffsetPropertyName = batchOffsetProperty
      ? batchOffsetProperty[0].name
      : null;
    const batchFilePropertyName = batchFileNameProperty
      ? batchFileNameProperty[0].name
      : null;
    let importLines = [];

    if (fileUrl)
      importLines = await getImportLines(
        fileUrl,
        fileType,
        csvDelimiter,
        logging,
      );
    else throwError('No import URL found.');
    if (logging) {
      console.log(`Lines in import file: ${importLines.length}`);
    }

    if (!batched && importLines.totalCount > 50000)
      throwError(
        'The number of import lines is too large (more than 50000). Please split your import file into several smaller files.',
      );

    if (validateRequiredColumns) {
      if (logging) {
        console.log('Validating required columns');
      }
      if (importLines.length === 0) {
        throwError(
          'There are no import lines found. The required column validation could not be done',
        );
      }
      const fileColumns = Object.keys(importLines[0]);
      const propertyMappingsRequiredKeys = propertyMappings
        .filter((propertyMap) => propertyMap.key.endsWith('*'))
        .map((propertyMap) => propertyMap.key);
      // Check if all required values are present
      const missingValues = propertyMappingsRequiredKeys.filter(
        (requiredValue) =>
          !fileColumns.includes(requiredValue) &&
          !fileColumns.includes(requiredValue.slice(0, -1)), // Remove the last character (the asterisk)
      );

      if (missingValues.length > 0) {
        // Throw an error if any required values are missing
        throwError(`Required values are missing: ${missingValues.join(', ')}`);
      }

      if (logging) {
        console.log('All required values are present');
      }
    }

    const propertyMappingMain = prepareRelationMappings(
      propertyMappings,
      propertyMappingsFormat,
    );

    if (logging) {
      console.log(
        'Main mappings (first 2000 characters): ' +
          JSON.stringify(propertyMappingMain),
      );
    }

    const propertyMappingsUpdateSpecific = prepareRelationMappings(
      propertyMappingsUpdate,
    );
    const propertyMappingsUpdateSpecificString = JSON.stringify(
      propertyMappingsUpdateSpecific,
    );
    if (logging && propertyMappingsUpdateSpecificString !== '[]') {
      console.log(
        'Update specific mappings (first 2000 characters): ' +
          propertyMappingsUpdateSpecificString.substring(0, 2000),
      );
    }

    if (batched) {
      let currentBatch = 0;
      importLinesToCreateInBatches = splitArray(importLines, batchSize);
      if (logging)
        console.log(
          `Number of batches: ${importLinesToCreateInBatches.length} (batch size:  ${batchSize})`,
        );

      const propertiesBatchDBNames = [];
      propertiesBatchDBNames.push(batchFilePropertyName);
      propertiesBatchDBNames.push(batchOffsetPropertyName);

      where = `{ ${batchFilePropertyName}: { eq: "${fileUrl}" }}`;

      let batchRecord = await getOne(
        batchModelName,
        propertiesBatchDBNames,
        where,
      );

      if (batchRecord) {
        currentBatch = batchRecord[batchOffsetPropertyName];
        if (logging)
          console.log(
            `Existing batch record for this file found, continuing with batch: ${
              currentBatch + 1
            }`,
          );
      } else {
        const newBatchRecord = {};
        newBatchRecord[batchFilePropertyName] = fileUrl;
        newBatchRecord[batchOffsetPropertyName] = 0;
        newBatchRecord[batchSizePropertyName] = 0;
        const createdBatchRecord = await createOne(
          batchModelName,
          newBatchRecord,
        );
        newBatchRecord.id = createdBatchRecord.id;
        batchRecord = newBatchRecord;
      }

      for (i = currentBatch; i < importLinesToCreateInBatches.length; i++) {
        if (logging) {
          console.log(
            `Starting batch ${i + 1}: (import lines ${i * batchSize} to ${
              i * batchSize + importLinesToCreateInBatches[i].length
            } )`,
          );
        }

        const result = await processImportLines(
          importLinesToCreateInBatches[i],
          propertyMappingMain,
          defaultMappings,
          propertyMappings,
          propertyMappingsFormat,
          propertyMappingsUpdateSpecific,
          deduplicate,
          uniqueRecordColumnName,
          uniqueRecordColumnType,
          batched,
          batchSize,
          i,
          now(),
          logging,
          modelName,
        );

        if (result) {
          batchRecord[batchOffsetPropertyName] = i + 1;
          batchRecordWithoutId = { ...batchRecord };
          delete batchRecordWithoutId.id;
          await updateOne(batchModelName, batchRecord.id, batchRecordWithoutId);
          if (batched && logging) {
            console.log(result);
          }
        }
      }

      await deleteOne(batchModelName, batchRecord.id);
    } else {
      await processImportLines(
        importLines,
        propertyMappingMain,
        defaultMappings,
        propertyMappings,
        propertyMappingsFormat,
        propertyMappingsUpdateSpecific,
        deduplicate,
        uniqueRecordColumnName,
        uniqueRecordColumnType,
        false,
        0,
        0,
        now(),
        logging,
        modelName,
      );
    }
    return {
      result: `Finished processing ${importLines.length} import lines.`,
    };
  } catch (error) {
    throwError(error);
  }
};

export default importFile;
