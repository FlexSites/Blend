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

    $('#blend').on('click', function () {
        if (!$(this).hasClass('disabled')) {
            var selectedYears = [].slice.call($('#years option:selected').map(function (y) {
                return parseInt($(this).val())
            }));
            var selectedOptions = [].slice.call($('#facilities option:selected').map(function (a, item) {
                return { id: item.value, title: item.innerHTML }
            }));
            var years = encodeURIComponent(JSON.stringify(selectedYears));
            var options = encodeURIComponent(JSON.stringify(selectedOptions));
            var sepTXPOC = $('#sep-txpoc').is(':checked');
            var sepColor = $('#sep-color').is(':checked');
            $('#select').hide();
            $('.progress').show();
            $('#status').html('Locating Excel files&hellip;');
            $.get('/api/blend?options=' + options + '&years=' + years + '&txpoc=' + sepTXPOC + '&color=' + sepColor, function(data) {

            });
        }
    })


    function displayFacilities (files) {
        var select = document.createElement('select')
        select.id = 'facilities'
        select.name = 'facilities'
        select.multiple = 'multiple'

        // filter out certain workbooks
        if (files && files.facilities.length) {
            files.facilities.forEach(function (file) {
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