# excel-to-json
[![Java CI with Maven](https://github.com/javadev/excel-to-json/actions/workflows/maven.yml/badge.svg)](https://github.com/javadev/excel-to-json/actions/workflows/maven.yml)

### Excel to JSON Converter

This application converts Excel files into JSON format. It provides flexibility in formatting dates, parsing percent values, including rows with no data, rendering output as pretty formatted JSON, and filling columns with null values.

#### Usage

```
java -jar excel-to-json.jar [-s <source>] [-df <dateFormat>] [--help] [--percent] [--empty] [--pretty] [--fillColumns]
```

#### Options

- `-s, --source <source>`: The source file path to convert into JSON.
- `-df, --dateFormat <dateFormat>`: The template to use for formatting dates into strings.
- `-?, --help`: Display this help text.
- `--percent`: Parse percent values as floats.
- `--empty`: Include rows with no data.
- `--pretty`: Render output as pretty formatted JSON.
- `--fillColumns`: Fill rows with null values until they all have the same size.

#### Example

```sh
java -jar excel-to-json.jar -s input.xlsx --pretty
```

This command converts the `input.xlsx` Excel file into pretty formatted JSON.
