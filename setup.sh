#!/bin/bash

# Development Environment Setup Script
# This script helps set up the SharePoint Framework development environment

echo "🚀 SharePoint External User Manager - Setup Script"
echo "=================================================="

# Check if Node.js is installed
if ! command -v node &> /dev/null; then
    echo "❌ Node.js is not installed. Please install Node.js 16.x or 18.x"
    echo "   Download from: https://nodejs.org/"
    exit 1
fi

# Check Node.js version
NODE_VERSION=$(node -v)
echo "📋 Current Node.js version: $NODE_VERSION"

# Check if the version is compatible
MAJOR_VERSION=$(echo $NODE_VERSION | cut -d'.' -f1 | sed 's/v//')
if [ "$MAJOR_VERSION" -lt 16 ] || [ "$MAJOR_VERSION" -gt 18 ]; then
    echo "⚠️  Warning: Node.js version $NODE_VERSION may not be compatible"
    echo "   SharePoint Framework 1.18.2 requires Node.js 16.x or 18.x"
    echo "   Consider using a version manager:"
    echo "   - Windows: nvm-windows"
    echo "   - macOS/Linux: nvm"
    echo ""
    echo "   Install compatible version:"
    echo "   nvm install 18.17.1"
    echo "   nvm use 18.17.1"
    echo ""
    read -p "Continue anyway? (y/N): " -n 1 -r
    echo
    if [[ ! $REPLY =~ ^[Yy]$ ]]; then
        exit 1
    fi
fi

echo "📦 Installing dependencies..."
npm install

if [ $? -eq 0 ]; then
    echo "✅ Dependencies installed successfully!"
else
    echo "❌ Failed to install dependencies"
    echo "   This might be due to Node.js version compatibility"
    echo "   Please ensure you're using Node.js 16.x or 18.x"
    exit 1
fi

echo ""
echo "🎉 Setup complete! You can now:"
echo "   • Run 'npm run serve' to start development server"
echo "   • Run 'npm run build' to build the solution"
echo "   • Run 'npm run package-solution' to create deployment package"
echo ""
echo "📖 For detailed development guide, see DEVELOPER_GUIDE.md"