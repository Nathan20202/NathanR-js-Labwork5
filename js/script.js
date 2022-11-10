function userForm() {

    let last_name = document.getElementById('LastName').value;
    let first_name = document.getElementById('FirstName').value;
    let email_address = document.getElementById('EmailAddress').value;
    let address = document.getElementById('Address').value;
    let city = document.getElementById('City').value;
    let province = document.getElementById('Province').value;
    let membership;

    if (document.getElementById('Premium').checked) {
        membership = document.getElementById('Premium').value;
    } else {
        if (document.getElementById('Standard').checked) {
            membership = document.getElementById('Standard').value;
        } else if (document.getElementById('Basic').checked) {
            membership = document.getElementById('Basic').value;
        }
    }

    if (document.getElementById("p-output") == null) {
        let p = document.createElement('p');
        p.id = "p-output";
        p.innerHTML = "<br>" + 'Full Name: ' +last_name + ' ' +  first_name
            + "<br>" + 'Email: ' + email_address +
            "<br>" + 'Address: ' + address + "<br>" + 'City: ' + city
            + "<br>" + 'Province: ' +province+ "<br>" + 'Membership: ' +membership;
        document.getElementById('output').appendChild(p);
    } else {
        document.getElementById('p-output').innerHTML= "<br>" + 'Full Name: ' +last_name + ' ' +  first_name
            + "<br>" + 'Email: ' + email_address +
            "<br>" + 'Address: ' + address + "<br>" + 'City: ' + city
            + "<br>" + 'Province: ' +province+ "<br>" + 'Membership: ' +membership;
    }
}

if (document.location.pathname.endsWith("excel.html")) {
    document.getElementById("PhoneNumber").addEventListener("keydown", function (event) { if (event.key === "Enter") MyExcelFuns() })
}

function MyExcelFuns() {
    console.log("Submit - Excel")

    if (ValidNumber('PhoneNumber')) {
        resetErrorsExcel()

        let numbers = inputStringToArrayNumber('PhoneNumber')
        let result

        if (document.getElementById("AutoSum").checked) {
            result = autoSum(numbers)
        } else if (document.getElementById("Average").checked) {
            result = average(numbers)
        } else if (document.getElementById("Max").checked) {
            result = max(numbers)
        } else if (document.getElementById("Min").checked) {
            result = min(numbers)
        } else {
            result = "Error"
        }
        document.getElementById("Result").value = result
    } else {
        displayErrorsExcel()
    }
}

function ValidNumber(phoneNumber) {

    let inputStr = document.getElementById(phoneNumber).value
    let inputArr = inputStr.split("")
    let regexArray = ['0','1','2','3','4','5','6','7','8','9',' ','.',',']

    inputArr.forEach(character => {
        if (!matchInArray(character, regexArray)) {
            return false
        }
    })

    let numbers = inputStringToArrayNumber(phoneNumber)
    return numbers.length > 0;
}

function ValidNumberDisplay(phoneNumber) {
    if (ValidNumber(phoneNumber)) {
        resetErrorsExcel()
    } else {
        displayErrorsExcel()
    }
}

function inputStringToArrayNumber(phoneNumber) {
    let numberStr = document.getElementById(phoneNumber).value
    let numberArr = numberStr.split(" ")
    let numbers = []

    numberArr.forEach(element => {
        if (element != null && element !== "") {
            let number = Number(element)
            if (!isNaN(number)) {
                numbers.push(number)
            }
        }
    })
    return numbers;
}

function matchInArray(string, array) {

    for (let i = 0; i < array.length; i++) {
        if (string.match(array[i])) {
            return true
        }
    }
    return false
}

function displayErrorsExcel() {

    document.getElementById("PhoneNumber").classList.add("error-input")
    document.getElementById("Result").value = "Wrong Input!"

}

function resetErrorsExcel() {

    document.getElementById("PhoneNumber").classList.remove("error-input")
}

function autoSum(array) {

    let sum = 0

    array.forEach(element => {
        sum += element
    });
    return sum
}

function average(array) {

    let sum = 0
    let length = 0

    array.forEach(element => {
        sum += element
        length++
    });
    return sum/length
}

function max(array) {

    return Math.max(...array)
}

function min(array) {

    return Math.min(...array)
}