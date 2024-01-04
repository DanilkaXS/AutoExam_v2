function orientationEditor() {
    const orientationVal = $("#docOrientation").val();
    if (orientationVal == "vertical") {
        $("#biletEditor").css('width', '600px')
        $("#biletEditor").css('height', '750px')
        $("#verticalIcon").css("padding-left", "5px")
        $("#verticalIcon").css("display", "block")
        $("#horizontalIcon").css("display", "none")
        $("#biletEditor").css("display", "inline-block")


    } else if (orientationVal == "horizontal") {
        $("#biletEditor").css('width', '750px')
        $("#biletEditor").css('height', '600px')
        $("#verticalIcon").css("display", "none")
        $("#horizontalIcon").css("display", "block")
        $("#horizontalIcon").css("padding-left", "5px")
        $("#biletEditor").css("display", "inline-block")

    } else {
        $("#biletEditor").css("display", "none")
        $("#verticalIcon").css("display", "none")
        $("#horizontalIcon").css("display", "none")
    }
}


function editTemplateQuestionEditor() {
    $("#questionFieldDocumentEditor").find('div').remove()
    var numberQuestion = $("#numberQuestionBilet").val()
    for (var i = 1; i <= numberQuestion; i++) {
        $("#questionFieldDocumentEditor").append("<div id=\'questionTemplate" + i + "\'><p> " + i + ") <select name=\"\" id=\"templateQuestionType" + i + "\" onchange='questionTypeEditor(" + i + ")'>\n" +
            "                    <option value=\"default\">Тип питання</option>\n" +
            "                    <option value=\"0\">Тестове питання</option>\n" +
            "                    <option value=\"1\">Відкрите питання</option>\n" +
            "                </select>\n" +
            "                </p></div>")
    }

}

function questionTypeEditor(id) {
    const type = $("#templateQuestionType" + id).val();
    $("#questionTemplate" + id).find("i").remove()

    if (type == 0) {
        $("#questionTemplate" + id).append('<i>a) Відповідь  б) Відповідь  в) Відповідь  е) Відповідь</i>')
    } else if (type == 1) {
        $("#questionTemplate" + id).append('<i>Поле для відповіді</i>')
    } else {

    }
}

function testQuestionGenerator() {
    $("#fieldTestQuestion").find('div').remove()
    const number = $("#numberTestQuestion").val();
    for (var i = 1; i <= number; i++) {
        $("#fieldTestQuestion").append("<div style=\" margin: 5px\">\n" +
            "                    <p>" + i + ") <input type=\"text\" id=\"testQuestion" + i + "Text\" placeholder=\"Питання\"> <input type=\"number\"\n" +
            "                                                                                                  id=\"numberAnswerQuestion" + i + "\"\n" +
            "                                                                                                  min=\"0\"\n" +
            "                                                                                                  placeholder=\"Кількість відовідей\" onchange='answerTestQuestionGenerator(" + i + ")'>\n" +
            "                    <div id=\"fieldAnswerQuestion" + i + "\">\n" +
            "                    </div>\n" +
            "                    </p>\n" +
            "                </div>")
    }
}

function answerTestQuestionGenerator(id) {
    $("#fieldAnswerQuestion" + id).find('p').remove()
    const number = $("#numberAnswerQuestion" + id).val();
    for (let i = 1; i <= number; i++) {
        $("#fieldAnswerQuestion" + id).append("<p>" + i + ") <input type=\"text\" id=\"answerQuestion" + id + "Text" + i + "\" placeholder=\"Відповідь\"><select\n" +
            "                                id=\"answerQuestion" + id + "Type" + i + "\" style=\"margin-left: 5px\">\n" +
            "                            <option value=\"0\">Неправильно</option>\n" +
            "                            <option value=\"1\">Правильно</option>\n" +
            "                        </select>\n" +
            "                        </p>")
    }
}

function openQuestionGenerator() {
    $("#fieldOpenQuestion").find('div').remove()
    const number = $("#numberOpenQuestion").val()
    for (var i = 1; i <= number; i++) {
        $("#fieldOpenQuestion").append("<div style=\"margin: 5px\">\n" +
            "                    <label for=\"openQuestion" + i + "Text\">" + i + ")</label>\n" +
            "                    <input type=\"text\" id=\"openQuestion" + i + "Text\" placeholder=\'Питання\'>\n" +
            "                </div>")
    }
}

function generateBilets() {

    var questionType = {};
    for (let i = 1; i <= $("#numberQuestionBilet").val(); i++) {
        questionType["templateQuestionType" + i] = $("#templateQuestionType" + i).val()
    }

    var openQuestionField = {}
    var openQuestions = $("#openQuestions").val().split("\n")
    var numberOpenQuestion = openQuestions.length - 1

    for (let i = 1; i <= numberOpenQuestion; i++) {
        openQuestionField["openQuestion" + i] = {id: i, text: openQuestions[i-1]}
    }
    var testQuestionField = {}
    for (let i = 1; i <= $("#numberTestQuestion").val(); i++) {
        testQuestionField["testQuestion" + i] = {
            id: i,
            text: $("#testQuestion" + i + "Text").val(),
            numberAnswerQuestion: $("#numberAnswerQuestion" + i).val(),
            answers: {}
        }
        var answers = {}
        for (let j = 1; j <= $("#numberAnswerQuestion" + i).val(); j++) {
            answers["answer" + j] = {
                text: $("#answerQuestion" + i + "Text" + j).val(),
                status: $("#answerQuestion" + i + "Type" + j).val()

            }
            testQuestionField["testQuestion" + i]["answers"] = answers
        }
    }
    const data = {
        id: '',
        document: {
            nameDocument: $("#nameDocumentGen").val(),
            orientation: $("#docOrientation").val(),
            head: "Міністерство освіти і науки, молоді та спорту України\n" +
                "Ізмаїльський державний гуманітарний університет",
            osvitniyStypin: $("#osvitniyStypin").val(),
            osvitnyaPrograma: $("#osvitnyaPrograma").val(),
            group: $("#group").val(),
            specialnost: $("#specialnost").val(),
            specializacia: $("#specializacia").val(),
            semestr: $("#semestr").val(),
            disciplina: $("#disciplina").val(),
            numberQuestionBilet: $("#numberQuestionBilet").val(),
            questionType: {},
            kafedra: $("#kafedra").val(),
            protokol: [$("#numberProtokol").val(), $("#dateNumber").val(), $("#month").val(), $("#year").val()],
            persKafedra: $("#persKafedra").val(),
            examinator: $("#examinator").val(),
            numberBilet: $("#numberBilet").val()
        },
        questions: {
            openQuestions: {
                numberOpenQuestions: numberOpenQuestion,
                questionsField: {}
            },
            testQuestions: {
                numberTestQuestions: $("#numberTestQuestion").val(),
                questionsField: {
                    testQuestion1: {
                        id: 1,
                        text: "text",
                        numberAnswerQuestion: 1,
                        answers: {
                            answer1: {
                                text: "answer1",
                                status: "1"
                            }
                        }
                    }
                }
            }
        }


    };
    data.questions.openQuestions.questionsField = openQuestionField
    data.questions.testQuestions.questionsField = testQuestionField
    data.document.questionType = questionType
    sendDataPOST(data)
}

function sendDataPOST(data) {
    const url = '/postrequest';

    let send = $.ajax({
        url: url,
        method: 'post',
        contentType: "application/json",
        dataType: 'json',
        data: JSON.stringify(data),
        async: false,
        success: function (result) {
            alert(result);
        }
    }).responseText;
    console.log(send)
    $("#body").append(`<a href='/download_file' ><button type='button' class='btn btn-primary' style='font-size: 20px'>Завантажити результат</button></a>`)


}





