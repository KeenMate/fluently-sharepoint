# Fluently Sharepoint
A library created for better experience with Microsoft Sharepoint CSOM library.

<div markdown="0" class="two-columns">
  <div markdown="1" class="column">
    using (var context = new ClientContext(SiteURL)) 
    {
      // do something
    }
  </div>
  <div markdown="1" class="column">
    var op = SiteURL.Create();
    op.SelectWeb("Dashboard");
  </div>
</div>

