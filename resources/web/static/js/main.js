require.config({
    baseUrl: 'js',
    shim: {
        "bootstrap" : ['jquery'],
        "bootstrapSelect" : ['jquery']
    },
    paths: {
        "async": './async/index',
        "jquery": 'jquery.min',
        "view": 'view',
        "bootstrap" : 'bootstrap.min',
        "bootstrapSelect" : 'bootstrap-multiselect'
    }
});

require(['jquery', 'view', 'bootstrap', 'bootstrapSelect'], function ($, view) {

$(function(){
    function loadFacilities () {
        $.get('/api/facilities', function(data) {
            displayFacilities(data);
        });
    }


    function displayFacilities (files) {
        var select = document.createElement('select')
        select.id = 'facilities'
        select.name = 'facilities'
        select.multiple = 'multiple'

        // filter out certain workbooks
        console.log(files);
        if (files && files.facilities.length) {
            files.facilities.forEach(function (file) {
                console.log(file);
                var option = document.createElement('option')
                option.value = file.value
                option.text = file.text
                select.appendChild(option)
            })
        }

        $('.progress').hide()
        $('#status').html('Choose facilities&hellip;')
        $('#select').prepend(select)
        $('#select').show()

        $('#facilities').multiselect({
            maxHeight: 310,
            buttonWidth: '400px',
            enableFiltering: true,
            testing: 'something',
            includeSelectAllOption: true,
            enableCaseInsensitiveFiltering: true,

            selectAllName: 'all-facilities',
            selectAllText: $(':input[name="all-facilities"]').prop('checked') ? 'Deselect All' : 'Select All',


            onChange: function (option, checked) {
                var selectedOptions = $('#facilities option:selected')

                if (selectedOptions.length > 0)
                    $('#blend').removeClass('disabled')
                else
                    $('#blend').addClass('disabled')
            },

            buttonText: function (options, select) {
                if (options.length === 0) {
                    return 'Select Facilities';
                }
                else if (options.length > 4) {
                    return options.length + ' selected'
                }
                else {
                    var labels = [];
                    options.each(function () {
                        if ($(this).attr('label') !== undefined) {
                            labels.push($(this).attr('label'))
                        } else {
                            labels.push($(this).html())
                        }
                    })

                    return labels.join(', ') + ''
                }
            }
        })
    }

    function init (view) {
        view.buildYears();
        loadFacilities ();
    }

    init(view);
});
});