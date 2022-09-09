(() => {
    function sheet2arr(sheet) {
        var result = [];
        var row;
        var rowNum;
        var colNum;
        var range = XLSX.utils.decode_range(sheet['!ref']);
        for (rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
            row = [];
            for (colNum = range.s.c; colNum <= range.e.c; colNum++) {
                var nextCell = sheet[
                    XLSX.utils.encode_cell({r: rowNum, c: colNum})
                    ];
                if (typeof nextCell === 'undefined') {
                    row.push(void 0);
                } else row.push(nextCell.w);
            }
            result.push(row);
        }

        return result;
    }

    async function parseFile(file) {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data);
        const sheetName = Object.keys(workbook.Sheets)[0];
        let rows = sheet2arr(workbook.Sheets[sheetName]);

        // remove header
        rows.shift();

        return rows;
    }

    function makeLocation(type, label, parentCode = '') {
        let code = _.snakeCase(label);

        return {
            code: parentCode ? `${parentCode}.${code}` : code,
            type,
            parent_code: parentCode,
            label,
        };
    }

    async function process(countryLabel, file) {
        let rows = await parseFile(file);

        let country = makeLocation('COUNTRY', countryLabel);

        let sqlValues = [makeSQLValues(country)];

        let stats = {
            district: 0,
            province: 0,
            ward: 0,
        };

        let pushedLocations = [];

        for (const row of rows) {
            let district = makeLocation('DISTRICT', row[0], country.code);
            let province = makeLocation('PROVINCE', row[1], district.code);
            let ward = makeLocation('WARD', row[2], province.code);

            if (!ward.label) continue;

            if (!pushedLocations.includes(district.code)) {
                sqlValues.push(makeSQLValues(district));
                pushedLocations.push(district.code);
                stats.district++;
            }

            if (!pushedLocations.includes(province.code)) {
                sqlValues.push(makeSQLValues(province));
                pushedLocations.push(province.code);
                stats.province++;
            }

            sqlValues.push(makeSQLValues(ward));
            stats.ward++;
        }

        return {
            stats,
            sql: "INSERT INTO `locations` (`code`, `type`, `parent_code`, `label`) VALUES \n" + sqlValues.join(",\n"),
        };
    }

    function makeSQLValues(location) {
        return `('${location.code}', '${location.type}', '${location.parent_code}', '${location.label}')`;
    }

    let loading = false;

    async function onSubmit(e) {
        e.preventDefault();

        if (loading) return;

        let countryLabel = document.getElementById('country').value;
        let file = document.getElementById('file').files[0];

        let button = document.getElementById('submit');
        let sql = document.getElementById('sql');
        let stats = document.getElementById('stats');

        loading = true;
        button.innerText = 'Loading...';
        stats.innerText = '';
        sql.value = '';

        try {
            let result = await process(countryLabel, file);
            stats.innerText = `Districts: ${result.stats.district} - Provinces: ${result.stats.province} - Wards: ${result.stats.ward}`;
            sql.value = result.sql;
        } catch (e) {
            alert(e.message);
        } finally {
            loading = false;
            button.innerText = 'Generate SQL';
        }
    }

    document.getElementById('form').addEventListener("submit", onSubmit, false);
})();