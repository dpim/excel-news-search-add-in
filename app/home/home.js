(function () {
  'use strict';
  var articlesArray = [];
  const tokenForCogServices = "Your-key-here"

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    jQuery(document).ready(function () {
      app.initialize();
      $('#lookup-button').click(getDataFromSelection);
      $('.insert-button').click(insertArticle);
      $('.source-button').click(goToSource);
    });
  };

  // Reads data from current document selection and displays a notification
  function getDataFromSelection() {
    if (Office.context.document.getSelectedDataAsync) {
      Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
        function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            var textVal = result.value;
            if (textVal.length > 0) {
              var url = "https://api.cognitive.microsoft.com/bing/v5.0/news/search?q=" + textVal + "&count=5&offset=0&mkt=en-us&safeSearch=Moderate"
              $.ajax({
                url: url,
                type: 'GET',
                dataType: 'json',
                headers: {
                  "Ocp-Apim-Subscription-Key": tokenForCogServices,
                  "Access-Control-Allow-Origin": "*"
                },
                success: function (data) {
                  var articles = [];
                  var i;
                  for (i = 0; i < data.value.length; i++) {
                    var headline = data.value[i].name;
                    var description = data.value[i].description;
                    var source = data.value[i].url;
                    articles.push(new ArticleItem(headline, description, source));
                  }
                  addArticlesToView(articles);
                  articlesArray = articles;
                }
              });
            }
          } else {
            app.showNotification('Error:', result.error.message);
          }
        }
      );
    } else {
      app.showNotification('Error:', 'Reading selection data not supported by host application.');
    }
  }

  function goToSource(i) {
    window.open(articlesArray[i].source);
  }

  function insertArticle(i) {
    Office.context.document.setSelectedDataAsync([[articlesArray[i].headline, articlesArray[i].description, articlesArray[i].source]], function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
      }
    });
  }

  function ArticleItem(headline, description, source) {
    this.headline = headline;
    this.description = description;
    this.source = source;
  }

  function addArticlesToView(articles) {
    var i;
    $("#article-container").html(""); //clear container 
    for (i = 0; i < articles.length; i++) {
      let j = i;

      var article = $("<div>");
      article.addClass("article-headline")
      article.html("<b class='article-headline'>" + articles[i].headline + "</b>");

      var article_description = $("<div>");
      article_description.html(articles[i].description);
      article_description.addClass("article-description")

      var sourceButton = $("<button>").text('Source');
      sourceButton.click(function () { goToSource(j) }).addClass('ms-Button');

      var insertButton = $("<button>").text('Insert');
      insertButton.click(function () { insertArticle(j) }).addClass('ms-Button');

      var article_source = $("<div>");
      article_source.append(sourceButton);
      article_source.append(insertButton);
      article_source.addClass("article-source")

      $("#article-container").append(article);
      $("#article-container").append(article_description);
      $("#article-container").append(article_source);
    }
  }
})();


