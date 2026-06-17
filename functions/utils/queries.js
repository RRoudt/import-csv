import { throwError } from "./helpers";

const getOne = async (modelName, properties, where) => {
  const queryName = `one${modelName}`;
  const query = `{
      ${queryName}(where: ${where}) {
        id
        ${properties.join("\n")}
      }
    }
  `;

  const { data, errors } = await gql(query);
  if (errors) {
    throw new Error(errors);
  }

  const { [queryName]: record } = data;

  return record;
};

const getAll = async (gqlQuery, skip, take, results = [], variables = {}) => {
  const gqlResponse = await gql(gqlQuery, { skip, take, ...variables });
  if (gqlResponse) {
    const gqlQueryObject = Object.values(gqlResponse)[0]; // the data object
    const tmpResults = Object.values(gqlQueryObject)[0]; // the all query object which contains the result and totalcount

    if (tmpResults.totalCount > 20000)
      throwError(
        "The number of records to update is too large. Please turn on batching for this step and select a batch size between 0 and 10.000."
      );

    skip += take;
    if (tmpResults.results.length) {
      const newResults = [...results, ...tmpResults.results];
      results = newResults;
      if (skip <= tmpResults.totalCount) {
        results = await getAll(gqlQuery, skip, take, results, variables);
      }
    }
  }
  return results;
};

export { getOne, getAll };
