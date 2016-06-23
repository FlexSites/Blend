require(['jquery', 'view', 'bootstrap'], function ($, view) {

$(function(){
	function loadFacilities () {
		$.get('/api/facilities', function(data) {
			console.log('sup', data);
		});
	}

	function displayFacilities () {

	}

	function init () {
		view.buildYears();
		loadFacilities ();
	}

	init();
});
});

