module.exports = function addPropertiesToTemplate(templateString, properties) {

    const template = JSON.parse(templateString);
    properties.forEach((element) => {
        template[element.name] = element.value;
    });
    return JSON.stringify(template);

}