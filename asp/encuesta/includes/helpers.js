	
	function setCollapsible(className) {
  	// Collapsible
        var coll = document.getElementsByClassName(className);
        var i;

        for (i = 0; i < coll.length; i++) {
         	coll[i].addEventListener("click", function() {
            this.classList.toggle("active");
            var content = this.nextElementSibling;
            if (!content.style.display || content.style.display === "none"){    
                content.style.display = "table-row";    					
            } else {      
                content.style.display = "none";      
            } 
        });
        }
    }
