# If you come from bash you might have to change your $PATH.
# export PATH=$HOME/bin:/usr/local/bin:$PATH

# Path to your oh-my-zsh installation.
export ZSH="/Users/zhangxinyuan/.oh-my-zsh"

# Set name of the theme to load --- if set to "random", it will
# load a random theme each time oh-my-zsh is loaded, in which case,
# to know which specific one was loaded, run: echo $RANDOM_THEME
# See https://github.com/robbyrussell/oh-my-zsh/wiki/Themes
ZSH_THEME="af-magic"

# Set list of themes to pick from when loading at random
# Setting this variable when ZSH_THEME=random will cause zsh to load
# a theme from this variable instead of looking in ~/.oh-my-zsh/themes/
# If set to an empty array, this variable will have no effect.
# ZSH_THEME_RANDOM_CANDIDATES=( "robbyrussell" "agnoster" )

# Uncomment the following line to use case-sensitive completion.
# CASE_SENSITIVE="true"

# Uncomment the following line to use hyphen-insensitive completion.
# Case-sensitive completion must be off. _ and - will be interchangeable.
# HYPHEN_INSENSITIVE="true"

# Uncomment the following line to disable bi-weekly auto-update checks.
# DISABLE_AUTO_UPDATE="true"

# Uncomment the following line to automatically update without prompting.
# DISABLE_UPDATE_PROMPT="true"

# Uncomment the following line to change how often to auto-update (in days).
# export UPDATE_ZSH_DAYS=13

# Uncomment the following line if pasting URLs and other text is messed up.
# DISABLE_MAGIC_FUNCTIONS=true

# Uncomment the following line to disable colors in ls.
# DISABLE_LS_COLORS="true"

# Uncomment the following line to disable auto-setting terminal title.
# DISABLE_AUTO_TITLE="true"

# Uncomment the following line to enable command auto-correction.
# ENABLE_CORRECTION="true"

# Uncomment the following line to display red dots whilst waiting for completion.
# COMPLETION_WAITING_DOTS="true"

# Uncomment the following line if you want to disable marking untracked files
# under VCS as dirty. This makes repository status check for large repositories
# much, much faster.
# DISABLE_UNTRACKED_FILES_DIRTY="true"

# Uncomment the following line if you want to change the command execution time
# stamp shown in the history command output.
# You can set one of the optional three formats:
# "mm/dd/yyyy"|"dd.mm.yyyy"|"yyyy-mm-dd"
# or set a custom format using the strftime function format specifications,
# see 'man strftime' for details.
# HIST_STAMPS="mm/dd/yyyy"

# Would you like to use another custom folder than $ZSH/custom?
# ZSH_CUSTOM=/path/to/new-custom-folder

# Which plugins would you like to load?
# Standard plugins can be found in ~/.oh-my-zsh/plugins/*
# Custom plugins may be added to ~/.oh-my-zsh/custom/plugins/
# Example format: plugins=(rails git textmate ruby lighthouse)
# Add wisely, as too many plugins slow down shell startup.
plugins=(git zsh-autosuggestions zsh-syntax-highlighting)

source $ZSH/oh-my-zsh.sh

# User configuration

# export MANPATH="/usr/local/man:$MANPATH"

# You may need to manually set your language environment
# export LANG=en_US.UTF-8

# Preferred editor for local and remote sessions
# if [[ -n $SSH_CONNECTION ]]; then
#   export EDITOR='vim'
# else
#   export EDITOR='mvim'
# fi

# Compilation flags
# export ARCHFLAGS="-arch x86_64"

# Set personal aliases, overriding those provided by oh-my-zsh libs,
# plugins, and themes. Aliases can be placed here, though oh-my-zsh
# users are encouraged to define aliases within the ZSH_CUSTOM folder.
# For a full list of active aliases, run `alias`.
#
set hlsearch
# Example aliases
# alias zshconfig="mate ~/.zshrc"
# alias ohmyzsh="mate ~/.oh-my-zsh"
#export HTTP_PROXY='http://localhost:8080'
#export HTTPS_PROXY='http://localhost:8080'
#export NO_PROXY="localhost,127.0.0.1,.oa.com"
#export http_proxy=$HTTP_PROXY
#export https_proxy=$HTTPS_PROXY
#export no_proxy=$NO_PROXY
export LC_ALL=en_US.UTF-8
export LANG=en_US.UTF-8
ctags=/usr/local/bin/ctags
alias work="~/workspace/golang/src && ls"
export GOROOT=/Users/zhangxinyuan/go
export GOPATH=$HOME/workspace/golang
export GOBIN=$GOPATH/bin
export PATH=$PATH:$GOROOT
#alias python="/Library/Frameworks/Python.framework/Versions/3.7/bin/python3.7"
alias updatedb='/usr/libexec/locate.updatedb'
alias spark='cd  /Users/zhangxinyuan/spark && ls'
alias info='cd /Users/zhangxinyuan/workspace/golang/src/gf/information && ls'
#export HOMEBREW_BOTTLE_DOMAIN=https://mirrors.aliyun.com/homebrew/homebrew-bottles
export HOMEBREW_NO_AUTO_UPDATE=true
JAVA_HOME=/Library/Java/JavaVirtualMachines/jdk1.8.0_201.jdk/Contents/Home
PATH=$JAVA_HOME/bin:$PATH:.
CLASSPATH=$JAVA_HOME/lib/tools.jar:$JAVA_HOME/lib/dt.jar:.
export JAVA_HOME
export PATH
export CLASSPATH

export HADOOP_HOME=/Library/hadoop-2.7.3
export PATH=$PATH:$HADOOP_HOME/bin

export SCALA_HOME=/Library/scala-2.11.8
export PATH=$PATH:$SCALA_HOME/bin
export SPARK_HOME=/Library/spark-2.3.0
export PATH=$PATH:$SPARK_HOME/bin
export HADOOP_COMMON_HOME=$HADOOP_HOME

export JAVA_LIBRARY_PATH=$HADOOP_HOME/lib/native/
export LD_LIBRARY_PATH=$HADOOP_HOME/lib/native/:$LD_LIBRARY_PATH
export SPARK_LOCAL_IP=127.0.0.1
export HADOOP_CONF_DIR=$HADOOP_HOME/etc/hadoop
export PATH=$PATH:$HADOOP_HOME/bin
export CGO_ENABLED=1
export PATH=$PATH:$GOPATH/bin:$GOROOT/bin
export PKG_CONFIG_PATH=/Users/zhangxinyuan/kafka/lib/pkgconfig
export ES_HOME=/Users/zhangxinyuan/spark/test/jadx/build/jadx/
export GETH=/Users/zhangxinyuan/workspace/golang/src/go-ethereum/build
export PATH=$PATH:$ES_HOME/bin
export PATH=$PATH:$GETH/bin
export HOMEBREW_BOTTLE_DOMAIN=https://mirrors.ustc.edu.cn/homebrew-bottles8

export INSTALL_LIB=/usr/local/lib
export LD_LIBRARY_PATH=/Users/zhangxinyuan/oracle/instantclient_19_8:$INSTALL_LIB
