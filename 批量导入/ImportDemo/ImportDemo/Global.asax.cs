using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.ComponentModel.Composition.Primitives;
using System.Web;
using System.Web.Http;
using System.Web.Http.Dependencies;
using System.Web.Mvc;
using System.Web.Routing;

namespace ImportDemo
{
    public class MvcApplication : System.Web.HttpApplication
    {
        private static MefDependencySolver solver = null;
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            //1.Mef接管
            DirectoryCatalog catalog = new DirectoryCatalog(AppDomain.CurrentDomain.SetupInformation.PrivateBinPath);
            solver = new MefDependencySolver(catalog);
            MefDependencySolver.Current = solver;
            DependencyResolver.SetResolver(solver);
            GlobalConfiguration.Configuration.DependencyResolver = solver;
            RouteConfig.RegisterRoutes(RouteTable.Routes);


        }
    }
    public class MefDependencySolver : System.Web.Mvc.IDependencyResolver, System.Web.Http.Dependencies.IDependencyResolver
    {
        public static MefDependencySolver Current { get; set; }

        //IDependencyResolver
        private readonly ComposablePartCatalog _catalog;
        private const string MefContainerKey = "MefContainerKey";

        public MefDependencySolver(ComposablePartCatalog catalog)
        {
            _catalog = catalog;
        }

        public CompositionContainer Container
        {
            get
            {
                CompositionContainer container;
                if (!HttpContext.Current.Items.Contains(MefContainerKey))
                {
                    container = new CompositionContainer(_catalog, CompositionOptions.DisableSilentRejection);
                    HttpContext.Current.Items.Add(MefContainerKey, container);
                    HttpContext.Current.DisposeOnPipelineCompleted(container);
                }
                else
                {
                    container = (CompositionContainer)HttpContext.Current.Items[MefContainerKey];
                }
                return container;
            }
        }

        #region IDependencyResolver Members

        public object GetService(Type serviceType)
        {
            string contractName = AttributedModelServices.GetContractName(serviceType);
            return Container.GetExportedValueOrDefault<object>(contractName);
        }

        public IEnumerable<object> GetServices(Type serviceType)
        {
            return Container.GetExportedValues<object>(serviceType.FullName);
        }

        public IDependencyScope BeginScope()
        {
            return this;
        }

        public void Dispose()
        {
        }
        #endregion
    }
}
