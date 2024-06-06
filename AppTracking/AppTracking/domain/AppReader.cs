class AppReader{

    private AppReaderImpl appReader;

    public AppReader(AppReaderImpl appReaderImpl){
        appReader = appReaderImpl;
    }

    public List<Dictionary<string, string>> getAppl(){
        return appReader.getApplications();
    }

};