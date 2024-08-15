

function copyCode() {
    navigator.clipboard.writeText(`
        function formatName(input) {
            return input
                .split(' ')         
                .map(word =>                   
                    word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()
                )
                .join(' ');                  
        };
        const opplist = document.querySelector(".opportunity-list");
        const personRows = document.querySelectorAll('.person-list-row');
        const personDataList = []; 

        personRows.forEach((personRow) => {
            const listitem1data = [];
            const data = personRow.querySelectorAll(".list-item-1");
            data.forEach((dataItem) => {
                listitem1data.push(dataItem.textContent.trim());
            });

            const name = personRow.querySelector('.person-name').textContent.trim();
            const id = listitem1data[0];
            const nationalityElements = personRow.querySelectorAll('.list-item .chip-container div');
            const homeLC = listitem1data[1];
            const homeMC = listitem1data[2];
            const opportunity = personRow.querySelector('.list-item-name .clickable-link').textContent.trim();
            const organization = listitem1data[3];
            const appliedAt = listitem1data[4];
            const phone = listitem1data[5];
            const nationality = Array.from(nationalityElements).map((el) => el.textContent.trim()).join(', ');    
            const pElement = document.querySelector('p[data-for="slots-data-cell"]');
            const dataTip = pElement.getAttribute('data-tip');
            const startDateRegex = /Start date: (\\d{2} \\w{3} \\d{4})/;
            const endDateRegex = /End date: (\\d{2} \\w{3} \\d{4})/;
            const startDateMatch = dataTip.match(startDateRegex);
            const endDateMatch = dataTip.match(endDateRegex);
            const personData = {
                "EP NAME": formatName(name),
                ID: id,
                NATIONALITY: nationality,
                "HOME LC": homeLC,
                "HOME MC": homeMC,
                "PROJECT NAME": opportunity,
                "HOST NAME": organization,
                "DATE OF APPLY": appliedAt,
                "PHONE NUMBER": String(phone),
                "START DATE": startDateMatch[1],
                "END DATE": endDateMatch[1],
            };
            personDataList.push(personData);
        });
            
            const sortColumnFlex = document.querySelector('.sort-column-flex');
            const buttonHTML = '<button id="copyButton" action="copyData()" type="button" style="padding: 10px 20px; background-color: #037EF3; color: white; border: none; border-radius: 10px; font-weight: 600; cursor: pointer; transition: background-color 0.3s ease;">Save Data</button>';
            sortColumnFlex.insertAdjacentHTML('beforeend', buttonHTML);
            document.getElementById("copyButton").addEventListener("click", copyData);

            var script = document.createElement('script');
            script.type = 'text/javascript';
            script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
            document.head.appendChild(script);
            function copyData() {
                const wb = XLSX.utils.book_new();
                const ws = XLSX.utils.json_to_sheet(personDataList);

                XLSX.utils.book_append_sheet(wb, ws, "Application Data");

                XLSX.writeFile(wb, "[EXPA]Application_data.xlsx");
            }
                
        `)
        .then(() => {
            alert('Đã sao chép đoạn script thành công');
        })
        .catch(err => {
            alert('Vui lòng báo lại cho chủ sở hữu. Lỗi sao chép: ', err);
        });
}

