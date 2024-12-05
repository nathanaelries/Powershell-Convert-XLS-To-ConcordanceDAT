# Configuration settings for Excel converter
$script:CONFIG = @{
    # Default characters
    PIPE = [char](20)
    THORNE = [char](254)
    TAB = [char](9)
    
    # Performance settings
    BATCH_SIZE = 10000           # Number of rows to process in memory at once
    MAX_THREADS = 4              # Maximum number of parallel processing threads
    BUFFER_SIZE = 65536         # File stream buffer size (64KB)
    READ_WRITE_BUFFER = 8192    # StreamReader/Writer buffer size (8KB)
}

# Export configuration
Export-ModuleMember -Variable CONFIG