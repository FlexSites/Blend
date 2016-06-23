define({
updateProgress : function (val) {
    $('.progress-bar').css('width', val + '%').attr('aria-valuenow', val)
},
resetProgress : function () {
    // turn off progress bar transitions, reset to 0, restore transitions
    $('.progress-bar').addClass('notransition')
    updateProgress(0)
    $('.progress-bar').remove('notransition')
},
buildYears : function () {
    //Refactor: Turn this into a function:
    var years = document.createElement('select')
    years.id = 'years'
    years.name = 'years'
    years.multiple = 'multiple'

    var thisYear = new Date().getFullYear()

    for (var y = 2011; y <= thisYear; y++) {
        var option = document.createElement('option')
        option.value = y
        option.text = y
        years.appendChild(option)
    }

    $('#select').prepend(years)

    $('#years').multiselect({
        includeSelectAllOption: true,
        buttonWidth: '160px',
        buttonText: function (options, select) {
            var len = select[0].options.length

            if (options.length === 0) {
                return 'Select Years'
            } else if (options.length === len) {
                return 'All Years'
            } else if (options.length > 2) {
                return options.length + ' years selected'
            } else {
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

    $('#years').multiselect('selectAll', false)
    $('#years').multiselect('updateButtonText')

    // move years over to make it and facilties into a combo button visually
    $('#years').next('.btn-group').css({ 'marginLeft': '-1px' })
},
resetForNew : function () {
    $('.multiselect-search').val('')
    $('#facilities').multiselect('deselectAll', false)
    $('#facilities').multiselect('updateButtonText')
    $('li').show().removeClass('filter-hidden')

    $('#years').multiselect('selectAll', false)
    $('#years').multiselect('updateButtonText')

    $('#quit').hide()
    $('#new').hide()

    $('#status').html('Choose facilities&hellip;')

    $('#blend').addClass('disabled')

    $('#select').show()
    $('#blend').show()
}

})